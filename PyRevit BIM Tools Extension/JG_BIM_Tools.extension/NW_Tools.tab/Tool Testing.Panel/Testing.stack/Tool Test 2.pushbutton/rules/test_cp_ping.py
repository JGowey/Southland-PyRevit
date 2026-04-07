# -*- coding: utf-8 -*-
RULE = {
    "name": "ZZZ TEST - CP_Ping",
    "description": "Simple test rule that writes 'PING' into Comments for pipes.",
    "priority": 5,
    "filter": {
        "categories": ["Pipes"],
        "parameter_rules": []
    },
    "target": {
        "name": "Comments",
        "guid": ""
    },
    "combine": "last_wins",
    "_prefer_shared": True,
    "batch": {"chunk_size": 500}
}

def predicate(element, context):
    # let the category filter do the heavy lifting
    return True

def compute(element, context):
    # always write this string
    return u"PING"
