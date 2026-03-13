"""
Main Speckle Automate function to generate all KPIs and Excel reports.
"""

from datetime import datetime
import os
import tempfile
from typing import Optional

import pandas as pd
from pydantic import Field, SecretStr
from speckle_automate import AutomateBase, AutomationContext, execute_automate_function

from excel_formatting import format_kpi_excel
from kpi_cfar import calculate_cfar
from kpi_mui import calculate_mui
from kpi_modularity import calculate_modularity_index
from kpi_energy import generate_energy_kpi_excel  # make sure this file is named exactly kpi_energy.py


class FunctionInputs(AutomateBase):
    whisper_message: SecretStr = Field(title="This is a secret message")
    forbidden_speckle_type: str = Field(
        title="Forbidden speckle type",
        description="If an object has this speckle_type, it will be marked with an error."
    )


def _write_kpi_excel(rows: list[dict], kpi_name: str) -> Optional[str]:
    """Write KPI table to Excel. Returns filepath or None if rows are empty."""
    if not rows:
        return None
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    filename = f"{kpi_name.lower().replace(' ', '_')}_{timestamp}.xlsx"
    filepath = os.path.join(tempfile.gettempdir(), filename)

    pd.DataFrame(rows).to_excel(filepath, index=False)
    format_kpi_excel(filepath, kpi_name)
    return filepath


def automate_function(automate_context: AutomationContext, function_inputs: FunctionInputs) -> None:
    # Receive the root object from Speckle
    version_root_object = automate_context.receive_version()

    # =========================
    # Generate KPIs
    # =========================
    cfar_file = _write_kpi_excel(calculate_cfar(version_root_object), "CFAR")
    mui_file = _write_kpi_excel(calculate_mui(version_root_object), "MUI")
    mod_file = _write_kpi_excel(calculate_modularity_index(version_root_object), "Modularity")

    # Energy KPI
    root_elements = getattr(version_root_object, "@elements", None) or getattr(version_root_object, "elements", [])
    facade = []
    for el in root_elements:
        name = getattr(el, "name", "").lower()
        objects = getattr(el, "@elements", None) or getattr(el, "elements", [])
        if "facade" in name:
            facade.extend(objects)

    energy_file = generate_energy_kpi_excel(facade)
    format_kpi_excel(energy_file, "Energy")

    # =========================
    # Upload all files
    # =========================
    files = {"CFAR": cfar_file, "MUI": mui_file, "Modularity": mod_file, "Energy": energy_file}
    download_urls = {}

    for kpi, file in files.items():
        if not file:
            continue
        blob_id = automate_context.store_file_result(file)
        file_name = os.path.basename(file)
        # Only blob ID is needed for reference
        download_urls[kpi] = f"Blob ID: {blob_id} (filename: {file_name})"

    # =========================
    # Report success
    # =========================
    message = "All KPIs generated successfully.\n"
    for kpi, url in download_urls.items():
        message += f"{kpi}: {url}\n"

    automate_context.mark_run_success(message)


if __name__ == "__main__":
    execute_automate_function(automate_function, FunctionInputs)