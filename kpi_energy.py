"""
kpi_energy_performance.py
Generates the Energy & Resource Performance KPI Excel sheet from a Speckle Facade collection.
Each row includes: Panel ID, Mesh ID, monthly ISR values, and annual total.
"""

import os
import tempfile
from datetime import datetime
import pandas as pd
from typing import List

def generate_energy_kpi_excel(facade_meshes: List) -> str:
    """
    Generates Excel sheet for Energy & Resource Performance KPI.

    Args:
        facade_meshes (list): List of facade mesh objects from Speckle, 
                              each having properties: 'panel_id', 'id', 'isr_value'
                              where 'isr_value' is a comma-separated string of 12 monthly values.

    Returns:
        str: Filepath of the generated Excel file.
    """

    rows = []

    # Sort meshes by panel_id
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
            "Annual Total": annual_total
        }

        rows.append(row)

    df = pd.DataFrame(rows)

    # Optional: Sort by Panel ID in Excel
    df.sort_values(by="Panel ID", inplace=True)

    # Create timestamped Excel file
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    filename = f"energy_performance_{timestamp}.xlsx"
    filepath = os.path.join(tempfile.gettempdir(), filename)

    df.to_excel(filepath, index=False)

    return filepath