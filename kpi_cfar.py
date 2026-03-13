# kpi_cfar.py

def calculate_cfar(version_root_object):
    """
    Calculate Column Free Area Ratio (CFAR) per level and for the entire building.

    Args:
        version_root_object: Root object received from Speckle Automate.

    Returns:
        List of dictionaries representing table rows for Excel.
    """

    # Get root collections
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

    # Containers for level aggregation
    slab_area_by_level = {}
    column_area_by_level = {}
    core_area_by_level = {}

    def add_area(obj, container, area_key):
        props = getattr(obj, "properties", None)
        if not props:
            return

        level = props.get("level")
        area = props.get(area_key)

        if level is None or area is None:
            return

        container[level] = container.get(level, 0) + area

    # Aggregate areas
    for obj in slabs:
        add_area(obj, slab_area_by_level, "slab_area")

    for obj in columns:
        add_area(obj, column_area_by_level, "column_area")

    for obj in cores:
        add_area(obj, core_area_by_level, "core_area")

    rows = []

    # Sort levels lowest → highest
    levels = sorted(slab_area_by_level.keys())

    total_slab = 0
    total_column = 0
    total_core = 0

    for level in levels:
        slab = slab_area_by_level.get(level, 0)
        column = column_area_by_level.get(level, 0)
        core = core_area_by_level.get(level, 0)

        column_free_area = slab - (column + core)
        cfar = column_free_area / slab if slab else 0

        rows.append(
            {
                "Level": level,
                "Slab Area": slab,
                "Column Area": column,
                "Core Area": core,
                "Column Free Area": column_free_area,
                "CFAR %": cfar * 100,
            }
        )

        total_slab += slab
        total_column += column
        total_core += core

    # Building totals
    building_free_area = total_slab - (total_column + total_core)
    building_cfar = building_free_area / total_slab if total_slab else 0

    building_row = {
        "Level": "Building Total",
        "Slab Area": total_slab,
        "Column Area": total_column,
        "Core Area": total_core,
        "Column Free Area": building_free_area,
        "CFAR %": building_cfar * 100,
    }

    # Put building row at the top
    rows.insert(0, building_row)

    return rows