"""
transfer_modularity_model.py
-----------------------------
Transfers only Facade and Exoskeleton to the modularity target model
with colour overrides based on normalised modularity value:
  - Facade:      blue (most repeated) → orange (least repeated)
  - Exoskeleton: green (most repeated) → red (least repeated)
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

FACADE_HIGH = (0,   120, 255)
FACADE_LOW  = (255, 140,   0)
EXO_HIGH    = (0,   200,  80)
EXO_LOW     = (220,  30,  30)


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

def _compute_normalised_values(facade_objects, exo_objects):
    facade_counts = defaultdict(int)
    exo_counts    = defaultdict(int)
    for obj in facade_objects:
        pid = get_prop(obj, "panel_id")
        if pid:
            facade_counts[str(pid).strip()] += 1
    for obj in exo_objects:
        mid = get_prop(obj, "member_id")
        if mid:
            exo_counts[str(mid).strip()] += 1
    total = len(facade_objects) + len(exo_objects)
    if total == 0:
        return {}, {}
    return (
        {k: v / total for k, v in facade_counts.items()},
        {k: v / total for k, v in exo_counts.items()},
    )


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

    # 1. Collect only Facade and Exoskeleton
    facade_objects = get_collection_objects(version_root, "Facade")
    exo_objects    = get_collection_objects(version_root, "Exoskeleton")
    print(f"[Modularity Transfer] Facade:{len(facade_objects)} Exo:{len(exo_objects)}")

    # 2. Compute normalised values
    facade_norm, exo_norm = _compute_normalised_values(facade_objects, exo_objects)

    # 3. Apply colours in-place
    for obj in facade_objects:
        pid    = str(get_prop(obj, "panel_id") or "").strip()
        colour = _colour_for_normalised(facade_norm.get(pid, 0.0), FACADE_HIGH, FACADE_LOW)
        _apply_colour_to_mesh(obj, colour)
    for obj in exo_objects:
        mid    = str(get_prop(obj, "member_id") or "").strip()
        colour = _colour_for_normalised(exo_norm.get(mid, 0.0), EXO_HIGH, EXO_LOW)
        _apply_colour_to_mesh(obj, colour)
    print(f"[Modularity Transfer] Colours applied to {len(facade_objects)+len(exo_objects)} objects.")

    # 4. Build new root with only Facade and Exoskeleton
    new_root = Base()
    new_root["speckle_type"] = "Speckle.Core.Models.Collections.Collection"
    new_root["name"]         = "Modularity Model"
    new_root["elements"]     = [
        _make_collection("Facade",      facade_objects),
        _make_collection("Exoskeleton", exo_objects),
    ]

    # 5. Send
    project_id = automate_context.automation_run_data.project_id
    print(f"[Modularity Transfer] Sending {len(facade_objects)+len(exo_objects)} objects...")

    try:
        transport = ServerTransport(stream_id=project_id, client=speckle_client)
        obj_id = operations.send(base=new_root, transports=[transport], use_default_cache=False)
        print(f"[Modularity Transfer] Sent: {obj_id}")
    except Exception as e:
        import traceback
        print(f"[Modularity Transfer] SEND ERROR: {type(e).__name__}: {e}")
        traceback.print_exc()
        raise

    # 6. Create version
    try:
        commit_id = speckle_client.commit.create(
            stream_id=project_id,
            object_id=obj_id,
            branch_name=target_branch,
            message="Automate: modularity model — colour coded by repetition index",
            source_application="SpeckleAutomate",
        )
        print(f"[Modularity Transfer] Version created: {commit_id}")
        return commit_id
    except Exception as e:
        import traceback
        print(f"[Modularity Transfer] COMMIT ERROR: {type(e).__name__}: {e}")
        traceback.print_exc()
        raise