"""
kpi_modularity.py
Calculates the Modularity Index KPI for Facade and Exoskeleton collections.
- Counts meshes per unique ID (panel_id for Facade, member_id for Exoskeleton)
- Calculates % of total elements per ID
- Building-level index: total unique IDs / total elements
"""

from typing import List, Dict
from datetime import datetime
import os
import tempfile
import pandas as pd


def calculate_modularity_index(version_root_object) -> str:
    """
    Calculate Modularity Index KPI and export to Excel.

    Args:
        version_root_object: Root Speckle object from Automate version.

    Returns:
        str: Filepath of the generated Excel file.
    """

    # Get root collections
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
    total_elements_count = len(all_elements)

    # Count meshes per unique ID
    facade_ids = {}
    for mesh in facade_elements:
        props = getattr(mesh, "properties", None)
        if props:
            pid = props.get("panel_id")
            if pid:
                facade_ids[pid] = facade_ids.get(pid, 0) + 1

    exo_ids = {}
    for mesh in exo_elements:
        props = getattr(mesh, "properties", None)
        if props:
            mid = props.get("member_id")
            if mid:
                exo_ids[mid] = exo_ids.get(mid, 0) + 1

    # Combine counts
    rows = []
    for pid, count in facade_ids.items():
        rows.append({
            "ID": pid,
            "Collection": "Facade",
            "Count of Meshes": count,
            "% of Total Elements": (count / total_elements_count) * 100
        })

    for mid, count in exo_ids.items():
        rows.append({
            "ID": mid,
            "Collection": "Exoskeleton",
            "Count of Meshes": count,
            "% of Total Elements": (count / total_elements_count) * 100
        })

    # Building-level index
    building_index = {
        "ID": "Building Total",
        "Collection": "-",
        "Count of Meshes": "-",
        "% of Total Elements": ((len(facade_ids) + len(exo_ids)) / total_elements_count) * 100
    }

    rows.insert(0, building_index)

    df = pd.DataFrame(rows)

    # Create timestamped Excel file
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    filename = f"modularity_index_{timestamp}.xlsx"
    filepath = os.path.join(tempfile.gettempdir(), filename)

    df.to_excel(filepath, index=False)

    return filepath