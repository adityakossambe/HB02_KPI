"""
Main Speckle Automate function to generate all KPIs in a single Excel workbook.
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
from kpi_energy import generate_energy_kpi_data  # updated function returns data


class FunctionInputs(AutomateBase):
    whisper_message: SecretStr = Field(title="This is a secret message")
    forbidden_speckle_type: str = Field(
        title="Forbidden speckle type",
        description="If an object has this speckle_type, it will be marked with an error."
    )


def _write_all_kpis_to_excel(kpi_data: dict[str, list[dict]]) -> str:
    """
    Writes all KPIs to a single Excel workbook with each KPI as a separate sheet.

    Args:
        kpi_data: dictionary where keys are sheet names, values are lists of row dictionaries.

    Returns:
        Path to the generated Excel file.
    """
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    filename = f"building_kpis_{timestamp}.xlsx"
    filepath = os.path.join(tempfile.gettempdir(), filename)

    with pd.ExcelWriter(filepath, engine="xlsxwriter") as writer:
        for sheet_name, rows in kpi_data.items():
            if not rows:
                continue
            df = pd.DataFrame(rows)
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            writer.sheets[sheet_name] = writer.book.get_worksheet_by_name(sheet_name)
            format_kpi_excel(filepath, sheet_name)

    return filepath


def automate_function(automate_context: AutomationContext, function_inputs: FunctionInputs) -> None:
    # Receive the root object from Speckle
    version_root_object = automate_context.receive_version()

    # =========================
    # Prepare KPI data
    # =========================
    kpi_data = {}

    kpi_data["CFAR"] = calculate_cfar(version_root_object)
    kpi_data["MUI"] = calculate_mui(version_root_object)
    kpi_data["Modularity"] = calculate_modularity_index(version_root_object)

    # Energy KPI
    root_elements = getattr(version_root_object, "@elements", None) or getattr(version_root_object, "elements", [])
    facade = []
    for el in root_elements:
        name = getattr(el, "name", "").lower()
        objects = getattr(el, "@elements", None) or getattr(el, "elements", [])
        if "facade" in name:
            facade.extend(objects)

    kpi_data["Energy"] = generate_energy_kpi_data(facade)

    # =========================
    # Write all KPIs to a single Excel file
    # =========================
    excel_file = _write_all_kpis_to_excel(kpi_data)

    # =========================
    # Upload file and report
    # =========================
    blob_id = automate_context.store_file_result(excel_file)
    file_name = os.path.basename(excel_file)
    download_message = f"Blob ID: {blob_id} (filename: {file_name})"

    message = f"All KPIs generated successfully in a single Excel file.\nDownload: {download_message}"
    automate_context.mark_run_success(message)


def automate_function_without_inputs(automate_context: AutomationContext) -> None:
    """A function example without inputs.

    If your function does not need any input variables,
     besides what the automation context provides,
     the inputs argument can be omitted.
    """
    pass


# make sure to call the function with the executor
if __name__ == "__main__":
    # NOTE: always pass in the automate function by its reference; do not invoke it!

    # Pass in the function reference with the inputs schema to the executor.
    execute_automate_function(automate_function, FunctionInputs)

    # If the function has no arguments, the executor can handle it like so
    # execute_automate_function(automate_function_without_inputs)
