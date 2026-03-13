# kpi_mui.py

def calculate_mui(version_root_object):
    """
    Calculate Material Usage Intensity (MUI) per level and for the entire building.

    MUI = (slab_volume + column_volume + core_volume) / slab_area

    Args:
        version_root_object: Root object received from Speckle Automate.

    Returns:
        List of dictionaries representing table rows for Excel.
    """

    root_elements = getattr(version_root_object, "@elements", None) or getattr(
        version_root_object, "elements", []
    )

    columns = []
    slabs = []
    cores = []

    # Identify collections
    for el in root_elements:
        name = getattr(el, "name", "").lower()
        objects = getattr(el, "@elements", None) or getattr(el, "elements", [])

        if "column" in name:
            columns.extend(objects)

        elif "slab" in name:
            slabs.extend(objects)

        elif "core" in name:
            cores.extend(objects)

    # Containers
    slab_area_by_level = {}
    slab_volume_by_level = {}
    column_volume_by_level = {}
    core_volume_by_level = {}

    def add_property(obj, container, key):
        props = getattr(obj, "properties", None)
        if not props:
            return

        level = props.get("level")
        value = props.get(key)

        if level is None or value is None:
            return

        container[level] = container.get(level, 0) + value

    # Aggregate data
    for obj in slabs:
        add_property(obj, slab_area_by_level, "slab_area")
        add_property(obj, slab_volume_by_level, "slab_volume")

    for obj in columns:
        add_property(obj, column_volume_by_level, "column_volume")

    for obj in cores:
        add_property(obj, core_volume_by_level, "core_volume")

    rows = []
    levels = sorted(slab_area_by_level.keys())

    total_slab_area = 0
    total_slab_volume = 0
    total_column_volume = 0
    total_core_volume = 0

    for level in levels:
        slab_area = slab_area_by_level.get(level, 0)
        slab_vol = slab_volume_by_level.get(level, 0)
        col_vol = column_volume_by_level.get(level, 0)
        core_vol = core_volume_by_level.get(level, 0)

        total_volume = slab_vol + col_vol + core_vol
        mui = total_volume / slab_area if slab_area else 0

        rows.append(
            {
                "Level": level,
                "Slab Area": slab_area,
                "Slab Volume": slab_vol,
                "Column Volume": col_vol,
                "Core Volume": core_vol,
                "Total Volume": total_volume,
                "MUI": mui,
            }
        )

        total_slab_area += slab_area
        total_slab_volume += slab_vol
        total_column_volume += col_vol
        total_core_volume += core_vol

    building_volume = total_slab_volume + total_column_volume + total_core_volume
    building_mui = building_volume / total_slab_area if total_slab_area else 0

    building_row = {
        "Level": "Building Total",
        "Slab Area": total_slab_area,
        "Slab Volume": total_slab_volume,
        "Column Volume": total_column_volume,
        "Core Volume": total_core_volume,
        "Total Volume": building_volume,
        "MUI": building_mui,
    }

    rows.insert(0, building_row)

    return rows