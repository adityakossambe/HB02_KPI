"""
_collection_helper.py
---------------------
Shared helper to extract a named sub-collection from the Grasshopper Model
root object and return its child elements as a flat list.

The Grasshopper Model structure is:
    root  (Base)
    ├── Slabs        (Base, children in elements / @elements)
    ├── Columns      (Base, children in elements / @elements)
    ├── Cores        (Base, children in elements / @elements)
    ├── Facade       (Base, children in elements / @elements)
    └── Exoskeleton  (Base, children in elements / @elements)
"""

from flatten import flatten_base


def get_collection_objects(root, collection_name: str) -> list:
    """
    Retrieve all mesh objects from a named sub-collection on the root.

    Tries both plain attribute and @-prefixed attribute access.
    Returns a flat list of all Base objects inside that collection,
    excluding the collection container itself.

    Args:
        root            : The version root object from receive_version()
        collection_name : e.g. "Slabs", "Facade", "Exoskeleton"

    Returns:
        List of Base objects (the individual meshes)
    """
    collection = getattr(root, collection_name, None)
    if collection is None:
        collection = getattr(root, f"@{collection_name}", None)
    if collection is None:
        return []

    # Flatten the collection, but exclude the collection container itself
    # (the last item yielded by flatten_base is the base object itself)
    all_objects = list(flatten_base(collection))

    # The collection root is the last item — drop it, keep only the children
    if all_objects and all_objects[-1] is collection:
        all_objects = all_objects[:-1]

    return all_objects


def get_prop(obj, *keys, default=None):
    """
    Safely retrieve a property from a Speckle Base object.
    Checks direct attributes first, then obj.properties, for each key.

    Args:
        obj     : Speckle Base object
        keys    : One or more property name strings to try in order
        default : Value to return if none of the keys are found

    Returns:
        The first matching value, or default.
    """
    for key in keys:
        val = getattr(obj, key, None)
        if val is None and hasattr(obj, "properties"):
            val = getattr(obj.properties, key, None)
        if val is not None:
            return val
    return default


def get_level(obj) -> str:
    """Extract the level name string from a Speckle object."""
    val = get_prop(obj, "level")
    if val is None:
        return "Unknown"
    if hasattr(val, "name"):
        return str(val.name)
    return str(val)