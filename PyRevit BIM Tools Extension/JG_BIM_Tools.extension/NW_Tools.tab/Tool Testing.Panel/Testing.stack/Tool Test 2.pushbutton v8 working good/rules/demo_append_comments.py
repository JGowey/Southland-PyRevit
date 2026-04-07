# -*- coding: utf-8 -*-
RULE = {
    "name": "Append to Comments",
    "description": "Appends '; Checked' to Comments",
    "priority": 60,                          # later = wins on same param
    "filter": { "categories": ["MEP Fabrication Pipework"] },
    "target": { "name": "Comments" },
    "combine": "append"
}

def compute(element, context):
    return u"Checked"
