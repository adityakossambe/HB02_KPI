"""This module contains the function's business logic.

Use the automation_context module to wrap your function in an Automate context helper.
"""

from pydantic import Field, SecretStr
from speckle_automate import (
    AutomateBase,
    AutomationContext,
    execute_automate_function,
)

from flatten import flatten_base


class FunctionInputs(AutomateBase):
    """These are function author-defined values.

    Automate will make sure to supply them matching the types specified here.
    Please use the pydantic model schema to define your inputs:
    https://docs.pydantic.dev/latest/usage/models/
    """

    # An example of how to use secret values.
    whisper_message: SecretStr = Field(title="This is a secret message")
    forbidden_speckle_type: str = Field(
        title="Forbidden speckle type",
        description=(
            "If a object has the following speckle_type,"
            " it will be marked with an error."
        ),
    )


def automate_function(
    automate_context: AutomationContext,
    function_inputs: FunctionInputs,
) -> None:
    """This is an example Speckle Automate function.

    Args:
        automate_context: A context-helper object that carries relevant information
            about the runtime context of this function.
            It gives access to the Speckle project data that triggered this run.
            It also has convenient methods for attaching results to the Speckle model.
        function_inputs: An instance object matching the defined schema.
    """
    import pandas as pd
    from datetime import datetime

    version_root_object = automate_context.receive_version()

    # Get root collections
    root_elements = getattr(version_root_object, "@elements", None) or getattr(
        version_root_object, "elements", []
    )

    cores = []
    columns = []
    slabs = []

    # Identify collections
    for el in root_elements:
        name = getattr(el, "name", "").lower()
        objects = getattr(el, "@elements", None) or getattr(el, "elements", [])

        if "core" in name:
            cores.extend(objects)

        elif "column" in name:
            columns.extend(objects)

        elif "slab" in name:
            slabs.extend(objects)

    # Dictionaries to aggregate area per level
    core_area_by_level = {}
    column_area_by_level = {}
    slab_area_by_level = {}

    def add_area(obj, container):
        props = getattr(obj, "properties", None)
        if not props:
            return

        area = props.get("Area")
        level = props.get("Level")

        if area is None or level is None:
            return

        container[level] = container.get(level, 0) + area

    for obj in cores:
        add_area(obj, core_area_by_level)

    for obj in columns:
        add_area(obj, column_area_by_level)

    for obj in slabs:
        add_area(obj, slab_area_by_level)

    # Determine roof level (highest slab level)
    roof_level = max(slab_area_by_level.keys()) if slab_area_by_level else None

    rows = []

    for level in sorted(slab_area_by_level.keys()):

        if level == roof_level:
            continue

        slab_area = slab_area_by_level.get(level, 0)
        core_area = core_area_by_level.get(level, 0)
        column_area = column_area_by_level.get(level, 0)

        service_area = core_area + column_area
        percent = (service_area / slab_area) * 100 if slab_area else 0

        rows.append(
            {
                "Level": level,
                "Slab Area": slab_area,
                "Core + Column Area": service_area,
                "Net Usable Area": slab_area - service_area,
                "Service %": percent,
            }
        )

    df = pd.DataFrame(rows)

    if not df.empty:
        avg_percent = df["Service %"].mean()

        avg_row = {
            "Level": "Average",
            "Slab Area": "",
            "Core + Column Area": "",
            "Net Usable Area": "",
            "Service %": avg_percent,
        }

        df = pd.concat([df, pd.DataFrame([avg_row])], ignore_index=True)

    # Create timestamped Excel file
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    filepath = f"/tmp/service_ratio_{timestamp}.xlsx"

    df.to_excel(filepath, index=False)

    file_url = automate_context.store_file_result(filepath)

    automate_context.mark_run_success(
        f"Core + Column vs Slab area analysis completed successfully.\n"
        f"Excel file uploaded: {file_url}"
    )   

    

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
