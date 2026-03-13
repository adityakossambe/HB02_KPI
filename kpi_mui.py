"""
kpi_mui.py
Calculates Material Usage Intensity (MUI) per level and for the entire building.
MUI = (slab_volume + column_volume + core_volume) / slab_area
"""

from typing import List, Dict

def calculate_mui(version_root_object) -> List[Dict]:
    """
    Calculate Material Usage Intensity (MUI) per level and building total.

    Args:
        version_root_object: Root Speckle object from Automate version.

    Returns:
        List of dictionaries representing table rows for Excel.
    """

    # Get root collections
    root_elements = getattr(version_root_object, "@elements", None) or getattr(
        version_root_object, "elements", []
    )

    columns, slabs, cores = [], [], []

    for el in root_elements:
        name = getattr(el, "name", "").lower()
        objects = getattr(el, "@elements", None) or getattr(el, "elements", [])

        if "column" in name:
            columns.extend(objects)
        elif "slab" in name:
            slabs.extend(objects)
        elif "core" in name:
            cores.extend(objects)

    # Containers per level
    slab_area_by_level = {}
    slab_volume_by_level = {}
    column_volume_by_level = {}
    core_volume_by_level = {}

    def add_property(obj, container, keys):
        props = getattr(obj, "properties", None)
        if not props:
            return

        prop_lower = {k.lower(): v for k, v in props.items()}

        level = prop_lower.get("level")
        value = None
        for key in keys:
            if key in prop_lower:
                value = prop_lower[key]
                break

        if level is None or value is None:
            return

        container[level] = container.get(level, 0) + value

    # Aggregate data
    for obj in slabs:
        add_property(obj, slab_area_by_level, ["slab_area"])
        add_property(obj, slab_volume_by_level, ["slab_volume"])

    for obj in columns:
        add_property(obj, column_volume_by_level, ["column_volume"])

    for obj in cores:
        add_property(obj, core_volume_by_level, ["core_volume"])

    rows = []
    levels = sorted(slab_area_by_level.keys())

    total_slab_area = 0
    total_slab_vol = 0
    total_column_vol = 0
    total_core_vol = 0

    for level in levels:
        slab_area = slab_area_by_level.get(level, 0)
        slab_vol = slab_volume_by_level.get(level, 0)
        col_vol = column_volume_by_level.get(level, 0)
        core_vol = core_volume_by_level.get(level, 0)

        total_vol = slab_vol + col_vol + core_vol
        mui = total_vol / slab_area if slab_area else 0

        rows.append({
            "Level": level,
            "Slab Area": slab_area,
            "Slab Volume": slab_vol,
            "Column Volume": col_vol,
            "Core Volume": core_vol,
            "Total Volume": total_vol,
            "MUI": mui,
        })

        total_slab_area += slab_area
        total_slab_vol += slab_vol
        total_column_vol += col_vol
        total_core_vol += core_vol

    # Building total
    building_total_vol = total_slab_vol + total_column_vol + total_core_vol
    building_mui = building_total_vol / total_slab_area if total_slab_area else 0

    building_row = {
        "Level": "Building Total",
        "Slab Area": total_slab_area,
        "Slab Volume": total_slab_vol,
        "Column Volume": total_column_vol,
        "Core Volume": total_core_vol,
        "Total Volume": building_total_vol,
        "MUI": building_mui,
    }

    # Insert building total at the top
    rows.insert(0, building_row)

    return rows