# -*- coding: utf-8 -*-
from __future__ import print_function, unicode_literals
import os, sys, traceback, datetime
from pyrevit import forms
from System import Action
from System.Windows.Threading import DispatcherPriority
from System.Threading import Thread, ThreadStart
from System.ComponentModel import INotifyPropertyChanged, PropertyChangedEventArgs
from System.Windows import Visibility

# -------- bundle shim --------
_BUNDLE_DIR = os.path.dirname(__file__)
if _BUNDLE_DIR not in sys.path:
    sys.path.insert(0, _BUNDLE_DIR)

THIS_DIR = os.path.dirname(__file__)
XAML_PATH = os.path.join(THIS_DIR, 'UI', 'DataSyncWindow.xaml')
LOG_PATH  = os.path.join(THIS_DIR, 'DataSync.log')

def _safe_str(x):
    try:
        if x is None: return u""
        if isinstance(x, unicode): return x
        return unicode(x)
    except:
        try: return unicode(str(x))
        except: return u"<unprintable>"

# -------- VM --------
class DataSyncVM(INotifyPropertyChanged):
    def __init__(self):
        self._handlers=[]
        self._IsLoading=True
        self._Status=u"Loading…"
        self._Total=0
        self._Done=0
        self._Error=None
        self._Rules=[]
        self._Groups=[]
        self._ScopeMode=u"activeview"
        self._Log=u""
    def add_PropertyChanged(self,h): self._handlers.append(h)
    def remove_PropertyChanged(self,h):
        try: self._handlers.remove(h)
        except: pass
    def _raise(self,n):
        a=PropertyChangedEventArgs(n)
        for h in list(self._handlers): h(self,a)
    def _set(self,f,v,n):
        if getattr(self,f)!=v: setattr(self,f,v); self._raise(n)
    @property
    def IsLoading(self): return self._IsLoading
    @IsLoading.setter
    def IsLoading(self,v): self._set('_IsLoading',bool(v),'IsLoading')
    @property
    def Status(self): return self._Status
    @Status.setter
    def Status(self,v): self._set('_Status',_safe_str(v),'Status')
    @property
    def Total(self): return self._Total
    @Total.setter
    def Total(self,v): self._set('_Total',int(v),'Total')
    @property
    def Done(self): return self._Done
    @Done.setter
    def Done(self,v): self._set('_Done',int(v),'Done')
    @property
    def Error(self): return self._Error
    @Error.setter
    def Error(self,v): self._set('_Error',_safe_str(v),'Error')
    @property
    def Rules(self): return self._Rules
    @Rules.setter
    def Rules(self,v): self._set('_Rules',list(v or []),'Rules')
    @property
    def Groups(self): return self._Groups
    @Groups.setter
    def Groups(self,v): self._set('_Groups',list(v or []),'Groups')
    @property
    def ScopeMode(self): return self._ScopeMode
    @ScopeMode.setter
    def ScopeMode(self,v): self._set('_ScopeMode',_safe_str(v),'ScopeMode')
    @property
    def Log(self): return self._Log
    @Log.setter
    def Log(self,v): self._set('_Log',_safe_str(v),'Log')

# -------- Window --------
class DataSyncWindow(forms.WPFWindow):
    _singleton=None

    @classmethod
    def show_singleton(cls):
        if cls._singleton and cls._singleton.IsVisible:
            try:
                cls._singleton.Activate()
                cls._singleton.Topmost = True; cls._singleton.Topmost = False
            except: pass
            return
        win = DataSyncWindow()
        cls._singleton = win
        win.Show()

    def __init__(self):
        forms.WPFWindow.__init__(self, XAML_PATH)
        self._run_inflight = False
        self._last_prog_ts = 0.0
        self._last_prog_pct = -1.0
        self.DataContext = DataSyncVM()
        self.Closed += self._on_closed

        # Progress row reserved; keep Hidden until run
        try: self.ProgressPanel.Visibility = Visibility.Hidden
        except: pass
        try: self.RadioActiveView.IsChecked = True
        except: pass

        # Model lists
        self._all_rules = []        # all discovered
        self._visible_rules = []    # after group filter
        self._display_rules = []    # after search filter (what the list shows)
        self._group_map = {}        # from groups json

        # Hook engine log into UI
        try:
            from engine import utils as _utils
            _utils.LOG_FN[0] = self._append_log
            self._safe_str = _utils.safe_str  # for search
        except:
            self._safe_str = _safe_str

        # Stage A load
        self.Dispatcher.BeginInvoke(DispatcherPriority.ContextIdle, Action(self._stage_a))

    def _on_closed(self, sender, e):
        DataSyncWindow._singleton = None
        try:
            from engine import utils as _utils
            _utils.LOG_FN[0] = None
        except: pass

    # ----- initial load -----
    def _stage_a(self):
        self._set_status(u"Loading rules and groups…")
        def work():
            try:
                from engine.utils import parse_rules_from_disk_fast, load_groups, load_groups_json, safe_str
                rules = parse_rules_from_disk_fast()
                for r in rules or []:
                    r['enabled'] = False
                    cats = r.get('categories') or []
                    r['categories_count']   = u"{} categories".format(len(cats)) if cats else u"0"
                    r['categories_tooltip'] = u", ".join([safe_str(c) for c in cats]) if cats else u""
                    r['basename'] = safe_str(os.path.basename(r.get('path',''))).lower()

                groups, gmap = load_groups_json()
                if not groups:
                    groups = load_groups(rules)
                    gmap = {}
                if u"All Rules" not in groups:
                    groups = [u"All Rules"] + list(groups)

                def ui():
                    vm=self.DataContext
                    self._all_rules = list(rules or [])
                    self._group_map = dict(gmap or {})
                    vm.Groups = list(groups or [])
                    try:
                        self.GroupsList.UnselectAll()
                        self.GroupsList.SelectedIndex = 0  # All Rules
                    except: pass
                    self._apply_group_filter_and_autocheck()
                    vm.IsLoading=False; self._set_status(u"Ready.")
                    try: self.ProgressPanel.Visibility = Visibility.Hidden
                    except: pass
                    self._append_log(u"Loaded {} rules; {} groups.".format(len(self._all_rules), len(vm.Groups)))
                self.Dispatcher.Invoke(Action(ui))
            except Exception as ex:
                def ui_err():
                    vm=self.DataContext
                    self._append_log(u"Failed to read rules/groups: {}\n{}".format(_safe_str(ex), _safe_str(traceback.format_exc())))
                    vm.IsLoading=False; self._set_status(u"No rules/groups found.")
                    try: self.ProgressPanel.Visibility = Visibility.Hidden
                    except: pass
                self.Dispatcher.Invoke(Action(ui_err))
        t=Thread(ThreadStart(work)); t.IsBackground=True; t.Start()

    # ----- filtering -----
    def _refresh_rules_view(self):
        vm = self.DataContext
        vm.Rules = []
        vm.Rules = list(self._display_rules)

    def _apply_search_filter(self):
        q = u""
        try:
            q = (self.RuleSearchBox.Text or u"").strip().lower()
        except: pass
        base = self._visible_rules
        if q:
            def match(r):
                if q in (_safe_str(r.get('name')).lower()): return True
                if q in (_safe_str(r.get('description')).lower()): return True
                cats = u", ".join([self._safe_str(c) for c in (r.get('categories') or [])]).lower()
                return q in cats
            self._display_rules = [r for r in base if match(r)]
        else:
            self._display_rules = list(base)
        self._refresh_rules_view()

    def _apply_group_filter_and_autocheck(self):
        try:
            selected = [unicode(x) for x in self.GroupsList.SelectedItems]
        except:
            selected = []
        show_all = (not selected) or any(g == u"All Rules" for g in selected)
        if show_all:
            self._visible_rules = list(self._all_rules)
            # keep default unchecked
            self._apply_search_filter()
            return

        if self._group_map:
            want=set()
            for g in selected: want |= set(self._group_map.get(g,set()))
            vis=[r for r in self._all_rules if _safe_str(r.get('basename')) in want]
        else:
            selset=set([_safe_str(g) for g in selected])
            vis=[r for r in self._all_rules if (_safe_str(r.get('group') or u"")) in selset]

        self._visible_rules = vis
        for r in self._visible_rules: r['enabled']=True
        self._apply_search_filter()

    # ----- run/stop -----
    def _start_scan(self):
        if getattr(self, "_run_inflight", False):
            try: self._append_log(u"Run already in progress; ignoring duplicate click.")
            except: pass
            return
        from engine.batched_loader import RevitBatchedLoader
        vm=self.DataContext

        rules_to_run = [r for r in self._display_rules if r.get('enabled')]
        if not rules_to_run:
            forms.alert(u"No rules are checked to run."); return

        ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self._append_log(u"\n" + u"="*82)
        self._append_log(u"Run started @ {}".format(ts))
        self._append_log(u"Scope: {}".format(u"Current Selection" if vm.ScopeMode==u"selection" else u"Active View"))
        self._append_log(u"Rules to run ({}): {}".format(
            len(rules_to_run), u", ".join([_safe_str(r.get('name')) for r in rules_to_run])))

        vm.IsLoading=True; vm.Done=0; vm.Total=0
        self._set_status(u"Scanning…")
        try: self.ProgressPanel.Visibility = Visibility.Visible
        except: pass

        chosen=[]
        try:
            for item in self.GroupsList.SelectedItems: chosen.append(_safe_str(item))
        except: pass
        scope={'mode': vm.ScopeMode, 'groups': chosen}

        def on_progress(done,total,msg=None):
            import time as _t
            # Always update final messages and phase messages
            if msg is not None:
                vm.Done, vm.Total = done, total
                self._append_log(_safe_str(msg))
                self._last_prog_ts = _t.time()
                self._last_prog_pct = (float(done)/float(total)) if total else 0.0
                return
            # Throttle frequent progress ticks (no message)
            now = _t.time()
            pct = (float(done)/float(total)) if total else 0.0
            emit = False
            if self._last_prog_pct < 0.0:
                emit = True
            elif abs(pct - self._last_prog_pct) >= 0.01:  # >=1%% change
                emit = True
            elif (now - self._last_prog_ts) >= 0.25:      # every 250ms
                emit = True
            if emit:
                vm.Done, vm.Total = done, total
                self._last_prog_ts = now
                self._last_prog_pct = pct

        def on_done(result):
            self._run_inflight = False
            vm.IsLoading=False
            try: self.ProgressPanel.Visibility = Visibility.Hidden
            except: pass
            if isinstance(result,dict) and result.get('error'):
                self._fatal(result['error']); return
            self._set_status(u"Ready.")
            self._append_log(u"Run finished.")
            self._append_log(u"Processed: {}/{} elements.".format(vm.Done,vm.Total))
            self._flush_log()

        RevitBatchedLoader.enqueue(on_progress,on_done,rules_to_run,scope)

    def _stop_scan(self):
        try:
            from engine.batched_loader import RevitBatchedLoader
            RevitBatchedLoader.cancel()
        except: pass
        self._append_log(u"Stop requested by user.")
        self._set_status(u"Stopped.")
        try: self.ProgressPanel.Visibility = Visibility.Hidden
        except: pass

    # ----- UI handlers -----
    def Run_Click(self,sender,e):  self._start_scan()
    def Stop_Click(self,sender,e): self._stop_scan()
    def Close_Click(self,sender,e): self._stop_scan(); self.Close()

    def GroupsList_SelectionChanged(self, sender, e): self._apply_group_filter_and_autocheck()
    def ScopeSelection_Checked(self,sender,e):
        if self.DataContext: self.DataContext.ScopeMode = u"selection"
    def ScopeActiveView_Checked(self,sender,e):
        if self.DataContext: self.DataContext.ScopeMode = u"activeview"

    def SelectAllGroups_Click(self,sender,e):
        try: self.GroupsList.SelectAll()
        except: pass
        self._apply_group_filter_and_autocheck()
    def ClearGroups_Click(self,sender,e):
        try:
            self.GroupsList.UnselectAll()
            self.GroupsList.SelectedIndex = 0
        except: pass
        self._apply_group_filter_and_autocheck()

    def RuleSelectAll_Click(self,sender,e):
        for r in self._display_rules: r['enabled'] = True
        self._refresh_rules_view()
    def RuleClear_Click(self,sender,e):
        for r in self._display_rules: r['enabled'] = False
        self._refresh_rules_view()
    def RuleSearchBox_TextChanged(self,sender,e):
        self._apply_search_filter()

    # ----- logging/helpers -----
    def _append_log(self,line):
        vm = self.DataContext
        at_bottom = False
        try:
            vo = self.LogTextbox.VerticalOffset
            vh = self.LogTextbox.ViewportHeight
            eh = self.LogTextbox.ExtentHeight
            at_bottom = (vo + vh) >= (eh - 2)
        except: at_bottom = True
        vm.Log = (vm.Log + u"\n" if vm.Log else u"") + _safe_str(line)
        try:
            if at_bottom: self.LogTextbox.ScrollToEnd()
        except: pass

    def _flush_log(self):
        try:
            with open(LOG_PATH,'wb') as f:
                f.write((self.DataContext.Log or u"").encode('utf-8'))
        except: pass

    def _fatal(self,message):
        vm=self.DataContext; vm.Error=message
        self._append_log(u"ERROR: "+_safe_str(message))
        self._flush_log()
        self._set_status(u"Failed. See log.")
        forms.alert(_safe_str(message))

    def _set_status(self,text):
        self.DataContext.Status=_safe_str(text)

if __name__=="__main__":
    try:
        DataSyncWindow.show_singleton()
    except Exception as ex:
        forms.alert(u"Failed to open DataSync UI:\n{}\n\n{}".format(
            _safe_str(ex), _safe_str(traceback.format_exc())))
