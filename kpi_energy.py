# kpi_energy.py
from typing import List, Dict

def generate_energy_kpi_data(facade_meshes) -> List[Dict]:
    """
    Returns data for Energy KPI (ISR values) without writing Excel.
    
    Args:
        facade_meshes: list of facade mesh objects with 'panel_id', 'id', 'isr_value'
    
    Returns:
        List of dictionaries representing table rows for Excel.
    """
    rows = []

    # Sort by panel_id
    sorted_meshes = sorted(facade_meshes, key=lambda m: getattr(m, "properties", {}).get("panel_id", ""))

    for mesh in sorted_meshes:
        props = getattr(mesh, "properties", None)
        if not props:
            continue

        panel_id = props.get("panel_id", "")
        mesh_id = props.get("id", "")
        isr_str = props.get("isr_value", "")

        if not isr_str:
            continue

        try:
            monthly_values = [float(v.strip()) for v in isr_str.split(",")]
        except Exception:
            continue

        if len(monthly_values) != 12:
            continue

        annual_total = sum(monthly_values)
        row = {
            "Panel ID": panel_id,
            "Mesh ID": mesh_id,
            "Jan": monthly_values[0],
            "Feb": monthly_values[1],
            "Mar": monthly_values[2],
            "Apr": monthly_values[3],
            "May": monthly_values[4],
            "Jun": monthly_values[5],
            "Jul": monthly_values[6],
            "Aug": monthly_values[7],
            "Sep": monthly_values[8],
            "Oct": monthly_values[9],
            "Nov": monthly_values[10],
            "Dec": monthly_values[11],
            "Annual Total": annual_total,
        }
        rows.append(row)

    return rows