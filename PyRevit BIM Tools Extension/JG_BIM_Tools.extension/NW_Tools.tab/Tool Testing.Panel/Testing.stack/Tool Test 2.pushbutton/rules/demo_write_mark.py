# -*- coding: utf-8 -*-
RULE = {
    "name": "Write DS mark",
    "description": "Sets Mark to DS-MARK (only if empty)",
    "priority": 40,
    "filter": { "categories": ["MEP Fabrication Pipework"] },
    "target": { "name": "Mark" },
    "combine": "only_if_empty"
}

def compute(element, context):
    return u"DS-MARK"
