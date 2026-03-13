"""
Main Speckle Automate function to generate all KPIs and formatted Excel reports.
"""

from datetime import datetime
import os
import tempfile
from pydantic import Field, SecretStr
from speckle_automate import AutomateBase, AutomationContext, execute_automate_function

from kpi_cfar import calculate_cfar
from kpi_mui import calculate_mui
from kpi_modularity import calculate_modularity_index
from kpi_energy import generate_energy_kpi_excel
from excel_formatting import format_kpi_excel


class FunctionInputs(AutomateBase):
    """User-defined inputs for the automation."""

    whisper_message: SecretStr = Field(title="This is a secret message")
    forbidden_speckle_type: str = Field(
        title="Forbidden Speckle type",
        description="If a Speckle object has this type, it will be marked as an error."
    )


def _write_kpi_excel(rows: list[dict], kpi_name: str) -> str:
    """Write a KPI table to a timestamped Excel file and return its path."""
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    filename = f"{kpi_name.lower().replace(' ', '_')}_{timestamp}.xlsx"
    filepath = os.path.join(tempfile.gettempdir(), filename)

    import pandas as pd
    pd.DataFrame(rows).to_excel(filepath, index=False)
    format_kpi_excel(filepath, kpi_name)

    return filepath


def automate_function(automate_context: AutomationContext, function_inputs: FunctionInputs) -> None:
    """Main automation function: generates all KPIs and uploads Excel reports."""

    # 1️⃣ Receive root version object from Speckle
    version_root_object = automate_context.receive_version()

    # 2️⃣ Generate CFAR KPI
    cfar_rows = calculate_cfar(version_root_object)
    cfar_file = _write_kpi_excel(cfar_rows, "CFAR")

    # 3️⃣ Generate MUI KPI
    mui_rows = calculate_mui(version_root_object)
    mui_file = _write_kpi_excel(mui_rows, "MUI")

    # 4️⃣ Generate Modularity KPI
    mod_file = calculate_modularity_index(version_root_object)
    format_kpi_excel(mod_file, "Modularity")

    # 5️⃣ Generate Energy KPI
    root_elements = getattr(version_root_object, "@elements", None) or getattr(version_root_object, "elements", [])
    facade_meshes = []
    for el in root_elements:
        if "facade" in getattr(el, "name", "").lower():
            facade_meshes.extend(getattr(el, "@elements", None) or getattr(el, "elements", []))

    energy_file = generate_energy_kpi_excel(facade_meshes)
    format_kpi_excel(energy_file, "Energy")

    # 6️⃣ Upload all Excel files and generate download URLs
    files = {"CFAR": cfar_file, "MUI": mui_file, "Modularity": mod_file, "Energy": energy_file}
    download_urls = {}
    for kpi, file_path in files.items():
        blob_id = automate_context.store_file_result(file_path)
        project_id = automate_context.project_id
        file_name = os.path.basename(file_path)
        download_urls[kpi] = f"https://speckle.xyz/projects/{project_id}/files/{blob_id}/{file_name}"

    # 7️⃣ Mark run success with all links
    message = "All KPIs generated successfully:\n"
    for kpi, url in download_urls.items():
        message += f"{kpi}: {url}\n"

    automate_context.mark_run_success(message)


def automate_function_without_inputs(automate_context: AutomationContext) -> None:
    """If no user inputs are required."""
    pass


if __name__ == "__main__":
    # Execute the automation function with the input schema
    execute_automate_function(automate_function, FunctionInputs)
    # If no inputs are needed, use:
    # execute_automate_function(automate_function_without_inputs)