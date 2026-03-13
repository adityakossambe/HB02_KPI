"""
transfer_modularity_model.py
-----------------------------
Transfers only Exoskeleton elements to the modularity target model
with green (most repeated) → red (least repeated) colour coding.
"""

from collections import defaultdict
from specklepy.objects.base import Base
from specklepy.transports.server import ServerTransport
from specklepy.api import operations

from _collection_helper import get_collection_objects, get_prop


# ── Colour helpers ────────────────────────────────────────────────────────

def _lerp(a, b, t):
    return int(a + (b - a) * t)

def _to_argb(r, g, b, a=255):
    return (a << 24) | (r << 16) | (g << 8) | b

EXO_HIGH = (0,   200,  80)   # green — most repeated
EXO_LOW  = (220,  30,  30)   # red   — least repeated


def _colour_for_normalised(normalised, high_rgb, low_rgb):
    t = max(0.0, min(1.0, normalised))
    return _to_argb(
        _lerp(low_rgb[0], high_rgb[0], t),
        _lerp(low_rgb[1], high_rgb[1], t),
        _lerp(low_rgb[2], high_rgb[2], t),
    )


def _apply_colour_to_mesh(obj, colour_int):
    try:
        vertices = getattr(obj, "vertices", None)
        vertex_count = max(1, len(vertices) // 3) if vertices else 1
    except Exception:
        vertex_count = 1
    obj["colors"] = [colour_int] * vertex_count


# ── Normalised values ─────────────────────────────────────────────────────

def _compute_exo_normalised(exo_objects):
    counts = defaultdict(int)
    for obj in exo_objects:
        mid = get_prop(obj, "member_id")
        if mid:
            counts[str(mid).strip()] += 1
    total = len(exo_objects)
    if total == 0:
        return {}
    return {k: v / total for k, v in counts.items()}


# ── Collection builder ────────────────────────────────────────────────────

def _make_collection(name, objects):
    col = Base()
    col["speckle_type"] = "Speckle.Core.Models.Collections.Collection"
    col["name"]         = name
    col["elements"]     = objects
    return col


# ── Main ──────────────────────────────────────────────────────────────────

def transfer_modularity_model(
    automate_context,
    speckle_client,
    version_root,
    target_stream_id,
    target_branch="main",
):
    print(f"[Modularity Transfer] Starting → model: {target_stream_id}")

    # 1. Collect Exoskeleton only
    exo_objects = get_collection_objects(version_root, "Exoskeleton")
    print(f"[Modularity Transfer] Exoskeleton: {len(exo_objects)} objects")

    # 2. Compute normalised values
    exo_norm = _compute_exo_normalised(exo_objects)

    # 3. Apply colours in-place
    for obj in exo_objects:
        mid    = str(get_prop(obj, "member_id") or "").strip()
        colour = _colour_for_normalised(exo_norm.get(mid, 0.0), EXO_HIGH, EXO_LOW)
        _apply_colour_to_mesh(obj, colour)
    print(f"[Modularity Transfer] Colours applied.")

    # 4. Build root with just Exoskeleton
    new_root = Base()
    new_root["speckle_type"] = "Speckle.Core.Models.Collections.Collection"
    new_root["name"]         = "Modularity Model"
    new_root["elements"]     = [
        _make_collection("Exoskeleton", exo_objects),
    ]

    # 5. Send
    project_id = automate_context.automation_run_data.project_id
    print(f"[Modularity Transfer] Sending {len(exo_objects)} objects to project {project_id}...")

    try:
        transport = ServerTransport(stream_id=project_id, client=speckle_client)
        obj_id = operations.send(base=new_root, transports=[transport], use_default_cache=False)
        print(f"[Modularity Transfer] Sent: {obj_id}")
    except Exception as e:
        import traceback
        print(f"[Modularity Transfer] SEND ERROR: {type(e).__name__}: {e}")
        traceback.print_exc()
        raise

    # 6. Commit to target model
    try:
        commit_id = speckle_client.commit.create(
            stream_id=project_id,
            object_id=obj_id,
            branch_name=target_branch,
            message="Automate: exoskeleton — colour coded by repetition index",
            source_application="SpeckleAutomate",
        )
        print(f"[Modularity Transfer] Committed: {commit_id}")
        return commit_id
    except Exception as e:
        import traceback
        print(f"[Modularity Transfer] COMMIT ERROR: {type(e).__name__}: {e}")
        traceback.print_exc()
        raise