"""
transfer_modularity_model.py
-----------------------------
Transfers the Exoskeleton collection as-is to the modularity target model.
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

    exo_objects = get_collection_objects(version_root, "Exoskeleton")
    print(f"[Modularity Transfer] Exoskeleton: {len(exo_objects)} objects")

    new_root = Base()
    new_root["speckle_type"] = "Speckle.Core.Models.Collections.Collection"
    new_root["name"]         = "Modularity Model"
    new_root["elements"]     = [_make_collection("Exoskeleton", exo_objects)]

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
        new_version_id = automate_context.create_new_version_in_project(
            root_object=new_root,
            model_id=target_stream_id,
            version_message="Automate: exoskeleton transfer",
        )
        print(f"[Modularity Transfer] Version created: {new_version_id}")
        return new_version_id
    except Exception as e:
        import traceback
        print(f"[Modularity Transfer] VERSION ERROR: {type(e).__name__}: {e}")
        traceback.print_exc()
        raise