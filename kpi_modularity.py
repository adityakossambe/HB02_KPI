# kpi_modularity.py
from typing import List, Dict

def calculate_modularity_index(version_root_object) -> List[Dict]:
    """
    Calculates Modularity Index KPI:
    - Collections included: Facade, Exoskeleton
    - Counts meshes per unique ID (panel_id/member_id)
    - Calculates percentage of total elements per ID
    - Top row shows building-level index (unique IDs / total elements)
    
    Returns:
        List of dictionaries representing rows for Excel.
    """
    root_elements = getattr(version_root_object, "@elements", None) or getattr(
        version_root_object, "elements", []
    )

    facade_elements = []
    exo_elements = []

    for el in root_elements:
        name = getattr(el, "name", "").lower()
        objects = getattr(el, "@elements", None) or getattr(el, "elements", [])
        if "facade" in name:
            facade_elements.extend(objects)
        elif "exoskeleton" in name:
            exo_elements.extend(objects)

    all_elements = facade_elements + exo_elements
    total_elements_count = len(all_elements) or 1  # avoid division by zero

    counts = []

    # Count Facade meshes by panel_id
    facade_ids = {}
    for mesh in facade_elements:
        props = getattr(mesh, "properties", None)
        if props and "panel_id" in props:
            pid = props["panel_id"]
            facade_ids[pid] = facade_ids.get(pid, 0) + 1

    # Count Exoskeleton meshes by member_id
    exo_ids = {}
    for mesh in exo_elements:
        props = getattr(mesh, "properties", None)
        if props and "member_id" in props:
            mid = props["member_id"]
            exo_ids[mid] = exo_ids.get(mid, 0) + 1

    # Combine counts
    for pid, count in facade_ids.items():
        counts.append({"ID": pid, "Collection": "Facade", "Count of Meshes": count})

    for mid, count in exo_ids.items():
        counts.append({"ID": mid, "Collection": "Exoskeleton", "Count of Meshes": count})

    # Add percentage of total elements
    for row in counts:
        row["% of Total Elements"] = (row["Count of Meshes"] / total_elements_count) * 100

    # Building-level index row
    building_index = {
        "ID": "Building Total",
        "Collection": "-",
        "Count of Meshes": "-",
        "% of Total Elements": (len(facade_ids) + len(exo_ids)) / total_elements_count * 100
    }

    # Return as list of dicts
    return [building_index] + counts