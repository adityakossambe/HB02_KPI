"""This module contains the function's business logic.

Use the automation_context module to wrap your function in an Automate context helper.
"""

import openpyxl

from speckle_automate import (
    AutomateBase,
    AutomationContext,
    execute_automate_function,
)

from kpi_cfar       import write_cfar_sheet
from kpi_mui        import write_mui_sheet
from kpi_modularity import write_modularity_sheet
from kpi_energy     import write_energy_sheet


class FunctionInputs(AutomateBase):
    """These are function author-defined values.

    Automate will make sure to supply them matching the types specified here.
    Additional inputs (e.g. target stream IDs for model transfers) will be
    added here when the transfer steps are implemented.
    """
    pass


def automate_function(
    automate_context: AutomationContext,
    function_inputs: FunctionInputs,
) -> None:
    """Speckle Automate function — generates a KPI Excel report from the
    triggering model version.

    Args:
        automate_context: Carries runtime context and provides access to the
            Speckle project data that triggered this run.
        function_inputs: An instance object matching the defined schema.
    """

    # ── 1. Receive the triggering version ────────────────────────────────
    version_root_object = automate_context.receive_version()

    if version_root_object is None:
        automate_context.mark_run_failed("Could not receive source model version.")
        return

    # ── 2. Build the Excel workbook with 4 KPI sheets ────────────────────
    try:
        wb = openpyxl.Workbook()
        wb.remove(wb.active)  # remove default empty sheet

        write_cfar_sheet(wb.create_sheet("CFAR"), version_root_object)
        write_mui_sheet(wb.create_sheet("MUI"), version_root_object)
        write_modularity_sheet(wb.create_sheet("Modularity Index"), version_root_object)
        write_energy_sheet(wb.create_sheet("Energy Performance"), version_root_object)

    except Exception as e:
        automate_context.mark_run_failed(f"Excel generation failed: {e}")
        raise

    # ── 3. Save and attach the Excel file to the run ──────────────────────
    excel_path = "/tmp/kpi_report.xlsx"
    wb.save(excel_path)
    automate_context.store_file_result(excel_path)

    # ── 4. Mark success ───────────────────────────────────────────────────
    automate_context.mark_run_success(
        "KPI report generated — 4 sheets: CFAR, MUI, Modularity Index, Energy Performance."
    )


if __name__ == "__main__":
    execute_automate_function(automate_function, FunctionInputs)