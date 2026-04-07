# Minimal rule template (safe to keep in repo as a starter)

RULE = {
    "name": "Demo Rule (TEST EDIT)",
    "description": "Stub: returns None (no writes).",
    "scope": "in_view",   # or "selection"
    "target": {"name": "CP_Demo", "guid": ""},  # not used yet
    "filter": {
        "categories": ["Pipes"],   # show how categories are declared
        "parameter_rules": []
    },
    "priority": 50,
    "combine": "last_wins",
    "batch": {"chunk_size": 500}
}

def compute(element, context):
    # Phase 1: no-op
    return None

# Optional extra narrowing (beyond JSON filter)
def predicate(element, context):
    return True
