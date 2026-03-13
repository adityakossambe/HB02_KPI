"""
_collection_helper.py
---------------------
Shared helper to extract a named sub-collection from the Grasshopper Model
root object and return its child elements as a flat list.
"""

from flatten import flatten_base


def get_collection_objects(root, collection_name: str) -> list:
    """
    Retrieve all mesh objects from a named sub-collection on the root.
    Tries multiple access patterns to handle different Speckle serialisation styles.
    """
    collection = None

    # Try direct attribute, @-prefixed, and dynamic prop lookup
    for key in [collection_name, f"@{collection_name}"]:
        collection = getattr(root, key, None)
        if collection is not None:
            break

    # Also try iterating root's dynamic attributes if nothing found yet
    if collection is None:
        try:
            for key in root.get_dynamic_member_names():
                if key.strip("@") == collection_name:
                    collection = root[key]
                    break
        except Exception:
            pass

    if collection is None:
        print(f"[WARN] Collection '{collection_name}' not found on root object.")
        print(f"[DEBUG] Root type: {type(root)}, speckle_type: {getattr(root, 'speckle_type', 'N/A')}")
        try:
            print(f"[DEBUG] Root dynamic members: {list(root.get_dynamic_member_names())}")
        except Exception:
            pass
        return []

    # Flatten the collection, exclude the collection container itself (last item)
    all_objects = list(flatten_base(collection))
    if all_objects and all_objects[-1] is collection:
        all_objects = all_objects[:-1]

    print(f"[INFO] Collection '{collection_name}': {len(all_objects)} objects found.")
    return all_objects


def get_prop(obj, *keys, default=None):
    """
    Safely retrieve a property from a Speckle Base object.
    Checks direct attributes first, then obj.properties, for each key.
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
                # properties may be a dict
                try:
                    val = props[key]
                except (KeyError, TypeError):
                    pass
            if val is not None:
                return val
        # Try dynamic member access
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