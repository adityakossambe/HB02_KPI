"""
Main Speckle Automate function to generate all KPIs and Excel reports.
"""

from datetime import datetime
import os
import tempfile

import pandas as pd
from pydantic import Field, SecretStr
from speckle_automate import (
    AutomateBase,
    AutomationContext,
    execute_automate_function,
)

from excel_formatting import format_kpi_excel
from kpi_cfar import calculate_cfar
from kpi_mui import calculate_mui
from kpi_modularity import calculate_modularity_index
from kpi_energy import generate_energy_kpi_excel


class FunctionInputs(AutomateBase):
    whisper_message: SecretStr = Field(title="This is a secret message")
    forbidden_speckle_type: str = Field(
        title="Forbidden speckle type",
        description="If an object has this speckle_type, it will be marked with an error."
    )


def _write_kpi_excel(rows: list[dict], kpi_name: str) -> str:
    """Write a KPI table to a timestamped Excel file and return its path."""
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    filename = f"{kpi_name.lower().replace(' ', '_')}_{timestamp}.xlsx"
    filepath = os.path.join(tempfile.gettempdir(), filename)

    pd.DataFrame(rows).to_excel(filepath, index=False)
    format_kpi_excel(filepath, kpi_name)
    return filepath


def automate_function(
    automate_context: AutomationContext,
    function_inputs: FunctionInputs,
) -> None:
    """Generates all KPIs and Excel reports for the building model."""
    
    version_root_object = automate_context.receive_version()

    # =========================
    # 1️⃣ CFAR KPI
    # =========================
    cfar_rows = calculate_cfar(version_root_object)
    cfar_file = _write_kpi_excel(cfar_rows, "CFAR")

    # =========================
    # 2️⃣ MUI KPI
    # =========================
    mui_rows = calculate_mui(version_root_object)
    mui_file = _write_kpi_excel(mui_rows, "MUI")

    # =========================
    # 3️⃣ Modularity KPI
    # =========================
    mod_file = calculate_modularity_index(version_root_object)
    format_kpi_excel(mod_file, "Modularity")

    # =========================
    # 4️⃣ Energy KPI
    # =========================
    root_elements = getattr(version_root_object, "@elements", None) or getattr(
        version_root_object, "elements", []
    )
    facade = []
    for el in root_elements:
        name = getattr(el, "name", "").lower()
        objects = getattr(el, "@elements", None) or getattr(el, "elements", [])
        if "facade" in name:
            facade.extend(objects)

    energy_file = generate_energy_kpi_excel(facade)
    format_kpi_excel(energy_file, "Energy")

    # =========================
    # Upload all Excel files
    # =========================
    files = {
        "CFAR": cfar_file,
        "MUI": mui_file,
        "Modularity": mod_file,
        "Energy": energy_file,
    }

    download_urls = {}
    for key, file in files.items():
        blob_id = automate_context.store_file_result(file)
        file_name = os.path.basename(file)
        # Use blob ID instead of project_id (fixed)
        download_urls[key] = f"Blob ID: {blob_id} (filename: {file_name})"

    # =========================
    # Mark run success with all links
    # =========================
    message = "All KPIs generated successfully.\n"
    for kpi, url in download_urls.items():
        message += f"{kpi}: {url}\n"

    automate_context.mark_run_success(message)


def automate_function_without_inputs(automate_context: AutomationContext) -> None:
    """Example function without inputs."""
    pass


if __name__ == "__main__":
    execute_automate_function(automate_function, FunctionInputs)