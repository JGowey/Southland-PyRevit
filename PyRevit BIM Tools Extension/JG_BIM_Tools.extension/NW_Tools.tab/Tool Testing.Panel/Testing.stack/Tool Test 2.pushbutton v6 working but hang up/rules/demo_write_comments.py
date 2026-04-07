# -*- coding: utf-8 -*-
RULE = {
    "name": "Write DS test to Comments",
    "description": "Writes 'DS Test' into Comments on fab pipework",
    "priority": 50,
    "filter": { "categories": ["MEP Fabrication Pipework"] },
    "target": { "name": "Comments" },
    "combine": "last_wins"
}

def compute(element, context):
    return u"DS Test"
