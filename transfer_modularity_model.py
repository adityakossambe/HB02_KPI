"""
transfer_modularity_model.py
-----------------------------
Transfers the Facade collection as-is to the modularity target model.
"""

from specklepy.objects.base import Base
from specklepy.transports.server import ServerTransport
from specklepy.api import operations

from _collection_helper import get_collection_objects


def _make_collection(name, objects):
    col = Base()
    col["speckle_type"] = "Speckle.Core.Models.Collections.Collection"
    col["name"]         = name
    col["elements"]     = objects
    return col


def transfer_modularity_model(
    automate_context,
    speckle_client,
    version_root,
    target_stream_id,
    target_branch="main",
):
    print(f"[Modularity Transfer] Starting → model: {target_stream_id}")

    facade_objects = get_collection_objects(version_root, "Facade")
    print(f"[Modularity Transfer] Facade: {len(facade_objects)} objects")

    new_root = Base()
    new_root["speckle_type"] = "Speckle.Core.Models.Collections.Collection"
    new_root["name"]         = "Modularity Model"
    new_root["elements"]     = [_make_collection("Facade", facade_objects)]

    project_id = automate_context.automation_run_data.project_id
    print(f"[Modularity Transfer] Sending...")

    try:
        transport = ServerTransport(stream_id=project_id, client=speckle_client)
        obj_id = operations.send(base=new_root, transports=[transport])
        print(f"[Modularity Transfer] Sent: {obj_id}")
    except Exception as e:
        import traceback
        print(f"[Modularity Transfer] SEND ERROR: {type(e).__name__}: {e}")
        traceback.print_exc()
        raise

    try:
        from specklepy.core.api.inputs.version_inputs import CreateVersionInput
        version_input = CreateVersionInput(
            project_id=project_id,
            model_id=target_stream_id,
            object_id=obj_id,
            message="Automate: facade transfer",
        )
        version = speckle_client.version.create(version_input)
        print(f"[Modularity Transfer] Version created: {version.id}")
        return version.id
    except Exception as e:
        import traceback
        print(f"[Modularity Transfer] COMMIT ERROR: {type(e).__name__}: {e}")
        traceback.print_exc()
        raise