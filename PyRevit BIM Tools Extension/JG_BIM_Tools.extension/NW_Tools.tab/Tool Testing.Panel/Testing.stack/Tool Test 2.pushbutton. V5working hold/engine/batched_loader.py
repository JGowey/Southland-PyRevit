# -*- coding: utf-8 -*-
from __future__ import unicode_literals

# Revit / pyRevit
from pyrevit import revit, DB, UI

# stdlib
import io, os, re, sys

# keep your original engine.utils intact for UI/rule discovery elsewhere
from engine.filters import compile_union_filter

# use only param_core for param work
from engine.param_core import (
    warm_type_param_cache_for_element,
    get_target_param,
    set_param_value,
    ensure_guid,
    safe_str,
    is_type_parameter_for_element,
    is_readonly,
    read_param_as_text,
)

BATCH_SIZE_SCAN  = 350
BATCH_SIZE_WRITE = 200

# -------- rule loading (works on IronPython) --------
def _as_module(ns, modname):
    class _M(object): pass
    m = _M()
    for k, v in ns.items():
        try: setattr(m, k, v)
        except: pass
    setattr(m, "__name__", modname)
    return m

def _load_rule_module(rule_dict):
    path = rule_dict.get('path')
    if not path or not os.path.isfile(path):
        return None

    rdir = os.path.dirname(path)
    if rdir and rdir not in sys.path:
        sys.path.insert(0, rdir)

    modname = "ds_rule_" + re.sub(r'[^0-9a-zA-Z_]+', '_', os.path.splitext(os.path.basename(path))[0])

    try:
        import imp
        return imp.load_source(modname, path)
    except Exception:
        pass

    ns = {'__file__': path, '__name__': modname, '__package__': None, '__builtins__': __builtins__}
    try:
        with io.open(path, 'r', encoding='utf-8-sig', errors='ignore') as f:
            src = f.read()
        if sys.version_info[0] >= 3:
            code = compile(src, path, 'exec'); exec(code, ns, ns)
        else:
            exec src in ns, ns
        return _as_module(ns, modname)
    except Exception:
        return None

# -------- selection normalization --------
def _get_selected_element_ids():
    ids = []
    # native selection (fastest)
    try:
        col = revit.uidoc.Selection.GetElementIds()
        if col and col.Count:
            for eid in col: ids.append(eid)
            return ids
    except Exception:
        pass
    # pyRevit wrapper fallback
    try:
        raw = list(revit.get_selection())
        for item in raw:
            try:
                if isinstance(item, DB.ElementId):
                    ids.append(item); continue
                if isinstance(item, DB.Reference):
                    ids.append(item.ElementId); continue
                if isinstance(item, DB.Element):
                    ids.append(item.Id); continue
                if isinstance(item, int):
                    ids.append(DB.ElementId(item)); continue
            except Exception:
                pass
    except Exception:
        pass
    return ids

# -------- combine policy (as in v1) --------
def _merge(existing, incoming):
    if existing is None:
        return incoming
    c = (incoming.get('combine') or 'last_wins').strip().lower()
    if c == 'first_wins':
        return existing
    if c == 'last_wins':
        if incoming.get('prio', 50) >= existing.get('prio', 50):
            return incoming
        return existing
    try:
        a = float(existing.get('value')); b = float(incoming.get('value'))
        if c == 'min': incoming['value'] = min(a, b); return incoming
        if c == 'max': incoming['value'] = max(a, b); return incoming
        if c == 'sum': incoming['value'] = a + b;   return incoming
    except Exception:
        pass
    if c in ('append','concat','join'):
        ea = safe_str(existing.get('value') or u'').strip()
        ib = safe_str(incoming.get('value') or u'').strip()
        parts = [p for p in (ea, ib) if p]
        incoming['value'] = u'; '.join(parts) if parts else u''
        return incoming
    if c == 'only_if_empty':
        ev = existing.get('value')
        if ev in (None, u'', 0, 0.0): return incoming
        return existing
    if c == 'skip_if_set':
        return existing
    return incoming

# -------- work state --------
class _WorkState(object):
    def __init__(self, rules, scope, on_progress):
        self.rules = rules or []
        self.scope = scope or {}
        self.on_progress = on_progress

        self._ids = None
        self._i = 0
        self.total = 0
        self._cancelled = False

        self._rule_mod = {}

        self.phase = 'scan'
        self.staged = {}         # (eid, pkey) -> update
        self.write_items = None
        self.write_idx = 0

        self.counts = {'elements':0, 'writes':0, 'conflicts':0, 'skips':0}
        self.fail   = {'missing':0, 'readonly':0, 'convert':0, 'typeparam':0, 'error':0}

        self._change_samples = []

    # -- prep
    def _prepare_scope(self):
        if self._ids is not None:
            return

        doc = revit.doc
        mode = (self.scope or {}).get('mode') or u'activeview'

        if mode == u'selection':
            ids = _get_selected_element_ids()
            scope_msg = u"Using Selection scope"
        elif mode == u'activeview':
            try:
                fec = DB.FilteredElementCollector(doc, doc.ActiveView.Id)
            except Exception:
                fec = DB.FilteredElementCollector(doc)
            fec = fec.WhereElementIsNotElementType().WherePasses(compile_union_filter(self.rules))
            ids = list(fec.ToElementIds()); scope_msg = u"Using Active View scope"
        else:
            fec = DB.FilteredElementCollector(doc).WhereElementIsNotElementType().WherePasses(compile_union_filter(self.rules))
            ids = list(fec.ToElementIds()); scope_msg = u"Using Entire Model scope"

        self._ids = ids
        self.total = len(ids)
        self._i = 0
        if self.on_progress:
            self.on_progress(0, self.total, u"{} ({} candidates).".format(scope_msg, self.total))

    def cancel(self): self._cancelled = True

    def _ensure_rule_module(self, rule):
        key = rule.get('path') or rule.get('name') or id(rule)
        if key in self._rule_mod: return self._rule_mod[key]
        mod = _load_rule_module(rule)
        self._rule_mod[key] = mod
        return mod

    # -- scan
    def _scan_batch(self):
        self._prepare_scope()
        if self._cancelled or self.total == 0:
            self.phase = 'apply'
            self.write_items = list(self.staged.items())
            self.write_idx = 0
            return

        doc = revit.doc
        hi = min(self._i + BATCH_SIZE_SCAN, self.total)

        for k in range(self._i, hi):
            if self._cancelled: break
            el = doc.GetElement(self._ids[k])
            if not el: continue

            for rule in (self.rules or []):
                mod = self._ensure_rule_module(rule)
                if not mod: continue
                pred = getattr(mod, 'predicate', None)
                compute = getattr(mod, 'compute', None)
                if compute is None: continue

                if pred is not None:
                    try:
                        if not pred(el, {'doc': doc}): continue
                    except Exception:
                        continue

                try:
                    val = compute(el, {'doc': doc})
                except Exception:
                    self.fail['error'] += 1
                    continue
                if val is None: continue

                tgt = rule.get('target') or {}
                tgt_name = safe_str(tgt.get('name') or u'')
                tgt_guid = ensure_guid(tgt.get('guid')) if tgt.get('guid') else None
                prefer_shared = rule.get('_prefer_shared')
                prio = rule.get('priority', 50)
                combine = (rule.get('combine') or 'last_wins').strip().lower()

                pkey = (safe_str(tgt_guid) if tgt_guid else safe_str(tgt_name))
                key = (el.Id.IntegerValue, pkey)

                incoming = {
                    'value': val, 'prio': prio, 'combine': combine,
                    'rule': rule.get('name', u'(rule)'),
                    'target_name': tgt_name, 'target_guid': tgt_guid, 'prefer_shared': prefer_shared
                }

                existing = self.staged.get(key)
                if existing is not None: self.counts['conflicts'] += 1
                self.staged[key] = _merge(existing, incoming)

        self._i = hi
        if self.on_progress: self.on_progress(self._i, self.total, None)

        if self._i >= self.total:
            self.phase = 'apply'
            self.write_items = list(self.staged.items())
            self.write_idx = 0
            self.counts['elements'] = len(set([eid for (eid, _p) in self.staged.keys()]))
            if self.on_progress:
                self.on_progress(0, len(self.write_items), u"Staged {} updates. Starting apply…".format(len(self.write_items)))

    # -- apply
    def _apply_batch(self):
        if self._cancelled: return
        doc = revit.doc
        total = len(self.write_items) if self.write_items else 0
        if self.write_idx >= total:
            self.phase = 'done'; return

        hi = min(self.write_idx + BATCH_SIZE_WRITE, total)
        tx = DB.Transaction(doc, 'DataSync: Apply')
        started = False

        try:
            tx.Start(); started = True
            for i in range(self.write_idx, hi):
                if self._cancelled:
                    try: tx.RollBack()
                    except: pass
                    return

                (eid_iv, _pkey), upd = self.write_items[i]
                el = doc.GetElement(DB.ElementId(int(eid_iv)))
                if not el:
                    self.counts['skips'] += 1; self.fail['missing'] += 1; continue

                # Optionally warm type cache only if we’ll touch type later (kept for symmetry)
                warm_type_param_cache_for_element(el)

                p = get_target_param(el,
                                     name=upd.get('target_name'),
                                     guid=upd.get('target_guid'),
                                     prefer_shared=upd.get('prefer_shared'))
                if not p:
                    self.counts['skips'] += 1; self.fail['missing'] += 1; continue

                if is_type_parameter_for_element(el, p):
                    self.counts['skips'] += 1; self.fail['typeparam'] += 1; continue

                if is_readonly(p):
                    self.counts['skips'] += 1; self.fail['readonly'] += 1; continue

                before = read_param_as_text(p)
                ok = set_param_value(p, upd['value'])
                if not ok:
                    self.counts['skips'] += 1; self.fail['convert'] += 1; continue
                after = read_param_as_text(p)

                if before == after:
                    self.counts['skips'] += 1
                else:
                    self.counts['writes'] += 1
                    if len(self._change_samples) < 10:
                        self._change_samples.append(
                            u"Element {} • {}: '{}' → '{}'".format(
                                el.Id.IntegerValue,
                                safe_str(p.Definition.Name if p.Definition else u'(param)'),
                                before, after
                            )
                        )

            tx.Commit()
            self.write_idx = hi
            if self.on_progress: self.on_progress(self.write_idx, total, None)

        except Exception:
            if started:
                try: tx.RollBack()
                except: pass
            raise

        if self.write_idx >= total:
            self.phase = 'done'
            if self.on_progress:
                summary = (u"Apply complete • elements: {e} • writes: {w} • conflicts: {c} • "
                           u"skips: {s} (missing {m}, readonly {ro}, convert {cv}, type {tp}, error {er})"
                           ).format(e=self.counts['elements'], w=self.counts['writes'], c=self.counts['conflicts'],
                                    s=self.counts['skips'], m=self.fail['missing'], ro=self.fail['readonly'],
                                    cv=self.fail['convert'], tp=self.fail['typeparam'], er=self.fail['error'])
                self.on_progress(self.write_idx, total, summary)
                for line in self._change_samples:
                    self.on_progress(self.write_idx, total, line)

    # -- stepper
    def step(self):
        if self.phase == 'scan':
            self._scan_batch()
            total = len(self.write_items or []) if self.phase == 'apply' else self.total
            done  = 0 if self.phase == 'apply' else self._i
            cont  = (self.phase != 'apply') or (total > 0)
            return (done, total, cont)
        elif self.phase == 'apply':
            self._apply_batch()
            total = len(self.write_items or [])
            return (self.write_idx, total, (self.phase != 'done'))
        else:
            return (self.write_idx, len(self.write_items or []), False)

# -------- ExternalEvent plumbing --------
class _BatchHandler(UI.IExternalEventHandler):
    def __init__(self):
        self._ev = None
        self._state = None
        self._on_prog = None
        self._on_done = None
    def Execute(self, app):
        try:
            done, total, cont = self._state.step()
            if self._on_prog: self._on_prog(done, total, None)
            if cont: self._ev.Raise()
            else:
                if self._on_done: self._on_done({'done': done, 'total': total})
        except Exception as ex:
            if self._on_done: self._on_done({'error': safe_str(ex)})
    def GetName(self): return "DataSync Batched Runner"

class RevitBatchedLoader(object):
    _handler = None
    _ev = None
    @classmethod
    def _ensure_event(cls):
        if cls._handler is None: cls._handler = _BatchHandler()
        if cls._ev is None: cls._ev = UI.ExternalEvent.Create(cls._handler)
        cls._handler._ev = cls._ev
    @classmethod
    def enqueue(cls, on_progress, on_done, rules, scope):
        cls._ensure_event()
        cls._handler._state = _WorkState(rules, scope, on_progress)
        cls._handler._on_prog = on_progress
        cls._handler._on_done = on_done
        cls._ev.Raise()
    @classmethod
    def cancel(cls):
        try:
            if cls._handler and cls._handler._state:
                cls._handler._state.cancel()
        except:
            pass
