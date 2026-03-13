"""
transfer_modularity_model.py
-----------------------------
Transfers Facade to the modularity target model.
Each mesh is coloured by its January ISR value using a
Ladybug-style blue → cyan → yellow → red heatmap.
"""

from specklepy.objects.base import Base
from specklepy.objects.other import RenderMaterial
from specklepy.transports.server import ServerTransport
from specklepy.api import operations

from _collection_helper import get_collection_objects, get_prop


# ── Ladybug-style heatmap ─────────────────────────────────────────────────
# Colour stops: blue → cyan → green → yellow → red
# t=0.0 → blue (lowest), t=1.0 → red (highest)

_STOPS = [
    (0.00, (74,  144, 226)),   # blue       (lowest)
    (0.20, (142, 195, 235)),   # light blue
    (0.40, (220, 240, 180)),   # light yellow-green
    (0.60, (255, 235,  80)),   # yellow
    (0.80, (255, 140,   0)),   # orange
    (1.00, (220,  30,   0)),   # red        (highest)
]


def _lerp_colour(t: float) -> int:
    """Map t (0.0–1.0) to an ARGB int using Ladybug colour stops."""
    t = max(0.0, min(1.0, t))
    for i in range(len(_STOPS) - 1):
        t0, c0 = _STOPS[i]
        t1, c1 = _STOPS[i + 1]
        if t <= t1:
            f = (t - t0) / (t1 - t0)
            r = int(c0[0] + (c1[0] - c0[0]) * f)
            g = int(c0[1] + (c1[1] - c0[1]) * f)
            b = int(c0[2] + (c1[2] - c0[2]) * f)
            return (0xFF << 24) | (r << 16) | (g << 8) | b
    r, g, b = _STOPS[-1][1]
    return (0xFF << 24) | (r << 16) | (g << 8) | b


def _get_jan_isr(obj) -> float | None:
    """Extract the January (first) ISR value from the isr_value property."""
    raw = get_prop(obj, "isr_value")
    if raw is None:
        return None
    try:
        first = str(raw).split(",")[0].strip()
        return float(first)
    except Exception:
        return None


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

    facade_objects = get_collection_objects(version_root, "Facade")
    print(f"[Modularity Transfer] Facade: {len(facade_objects)} objects")

    # 1. Extract January ISR values
    jan_values = []
    for obj in facade_objects:
        v = _get_jan_isr(obj)
        if v is not None:
            jan_values.append(v)

    if not jan_values:
        print("[Modularity Transfer] WARNING: No ISR values found, using red for all.")
        min_val = max_val = 0.0
    else:
        min_val = min(jan_values)
        max_val = max(jan_values)
        print(f"[Modularity Transfer] Jan ISR range: {min_val} → {max_val}")

    # 2. Apply colour per mesh
    val_range = max_val - min_val if max_val != min_val else 1.0
    for obj in facade_objects:
        v = _get_jan_isr(obj)
        t = (v - min_val) / val_range if v is not None else 0.0
        colour = _lerp_colour(t)
        obj["renderMaterial"] = RenderMaterial(
            name=f"ISR_Jan_{round(v or 0, 1)}",
            diffuse=colour,
            opacity=1.0,
        )

    print("[Modularity Transfer] Colours applied.")

    # 3. Build root and send
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
        new_version_id = automate_context.create_new_version_in_project(
            root_object=new_root,
            model_id=target_stream_id,
            version_message="Automate: facade — January ISR heatmap",
        )
        print(f"[Modularity Transfer] Version created: {new_version_id}")
        return new_version_id
    except Exception as e:
        import traceback
        print(f"[Modularity Transfer] VERSION ERROR: {type(e).__name__}: {e}")
        traceback.print_exc()
        raise