"""
_collection_helper.py
---------------------
Shared helper to extract a named sub-collection from the Grasshopper Model
root object and return its child elements as a flat list.

The root is a RootCollection. Named sub-collections (Slabs, Columns, etc.)
are children inside root.elements, each with a `name` property.
"""

from flatten import flatten_base


def get_collection_objects(root, collection_name: str) -> list:
    """
    Find a named sub-collection inside root.elements and return
    all its child mesh objects as a flat list.
    """
    # Get the top-level elements list from root
    root_elements = getattr(root, "elements", None) or getattr(root, "@elements", None)

    if root_elements is None:
        print(f"[WARN] Root has no 'elements' — cannot find '{collection_name}'.")
        return []

    # Find the matching collection by name
    target_collection = None
    for element in root_elements:
        element_name = getattr(element, "name", None)
        if element_name and element_name.strip() == collection_name:
            target_collection = element
            break

    if target_collection is None:
        print(f"[WARN] Collection '{collection_name}' not found in root.elements.")
        available = [getattr(e, "name", "?") for e in root_elements]
        print(f"[DEBUG] Available collections: {available}")
        return []

    # Flatten the collection, exclude the collection container itself (last item)
    all_objects = list(flatten_base(target_collection))
    if all_objects and all_objects[-1] is target_collection:
        all_objects = all_objects[:-1]

    print(f"[INFO] Collection '{collection_name}': {len(all_objects)} objects found.")
    return all_objects


def get_prop(obj, *keys, default=None):
    """
    Safely retrieve a property from a Speckle Base object.
    Checks direct attributes, then obj.properties (dict or object), for each key.
    """
    for key in keys:
        # Direct attribute
        val = getattr(obj, key, None)
        if val is not None:
            return val
        # Nested under .properties
        props = getattr(obj, "properties", None)
        if props is not None:
            val = getattr(props, key, None)
            if val is None:
                try:
                    val = props[key]
                except (KeyError, TypeError):
                    pass
            if val is not None:
                return val
        # Dynamic member access
        try:
            val = obj[key]
            if val is not None:
                return val
        except Exception:
            pass
    return default


def get_level(obj) -> str:
    """Extract the level name string from a Speckle object."""
    val = get_prop(obj, "level")
    if val is None:
        return "Unknown"
    if hasattr(val, "name"):
        return str(val.name)
    return str(val)


import re


def level_sort_key(s):
    """Sort levels numerically — levels are floats like -44.5, 0.0, 90.5."""
    try:
        return float(s)
    except (ValueError, TypeError):
        return 0.0


def id_sort_key(s):
    """
    Sort IDs by letter prefix then number suffix.
    e.g. LI1 < LI2 < LI10 < NQ1 < Q1 < Q2 < Q10 < T1 < T9 < T10
    """
    m = re.match(r'^([A-Za-z]+)(\d+)$', str(s).strip())
    if m:
        return (m.group(1).upper(), int(m.group(2)))
    return (str(s).upper(), 0)