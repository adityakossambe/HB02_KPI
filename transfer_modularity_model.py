"""
transfer_modularity_model.py
-----------------------------
Transfers Facade and Exoskeleton separately to the modularity target model.
No colour changes — plain transfer.
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


def _send(automate_context, speckle_client, name, objects, target_stream_id, message):
    project_id = automate_context.automation_run_data.project_id
    root = Base()
    root["speckle_type"] = "Speckle.Core.Models.Collections.Collection"
    root["name"]         = name
    root["elements"]     = [_make_collection(name, objects)]

    print(f"[Modularity Transfer] Sending {name} ({len(objects)} objects)...")
    transport = ServerTransport(stream_id=project_id, client=speckle_client)
    obj_id = operations.send(base=root, transports=[transport])
    print(f"[Modularity Transfer] {name} sent: {obj_id}")

    version_id = automate_context.create_new_version_in_project(
        root_object=root,
        model_id=target_stream_id,
        version_message=message,
    )
    print(f"[Modularity Transfer] {name} version: {version_id}")
    return version_id


def transfer_modularity_model(
    automate_context,
    speckle_client,
    version_root,
    target_stream_id,
    target_branch="main",
):
    print(f"[Modularity Transfer] Starting → model: {target_stream_id}")

    facade_objects = get_collection_objects(version_root, "Facade")
    exo_objects    = get_collection_objects(version_root, "Exoskeleton")
    print(f"[Modularity Transfer] Facade:{len(facade_objects)} Exo:{len(exo_objects)}")

    _send(automate_context, speckle_client, "Facade", facade_objects, target_stream_id,
          "Automate: facade transfer")

    _send(automate_context, speckle_client, "Exoskeleton", exo_objects, target_stream_id,
          "Automate: exoskeleton transfer")

    print("[Modularity Transfer] Done.")