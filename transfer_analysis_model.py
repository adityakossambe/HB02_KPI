"""
transfer_analysis_model.py
-----------------------------
Transfers Facade (with ISR heatmap colours) and Slabs (no colours)
to the analysis target model.
"""

from specklepy.objects.base import Base
from specklepy.objects.other import RenderMaterial
from specklepy.transports.server import ServerTransport
from specklepy.api import operations

from _collection_helper import get_collection_objects, get_prop


# ── Ladybug-style heatmap: blue → light blue → yellow → orange → red ─────

_STOPS = [
    (0.00, (74,  144, 226)),   # blue
    (0.20, (142, 195, 235)),   # light blue
    (0.40, (220, 240, 180)),   # light yellow-green
    (0.60, (255, 235,  80)),   # yellow
    (0.80, (255, 140,   0)),   # orange
    (1.00, (220,  30,   0)),   # red
]

MONTH_INDEX = {
    "January": 0, "February": 1, "March": 2, "April": 3,
    "May": 4, "June": 5, "July": 6, "August": 7,
    "September": 8, "October": 9, "November": 10, "December": 11,
}


def _lerp_colour(t: float) -> int:
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


def _get_isr_value(obj, isr_month: str) -> float | None:
    """Extract ISR value for the selected month or annual sum."""
    raw = get_prop(obj, "isr_value")
    if raw is None:
        return None
    try:
        values = [float(v.strip()) for v in str(raw).split(",")]
        if isr_month == "Annual":
            return sum(values)
        idx = MONTH_INDEX.get(isr_month, 0)
        return values[idx] if idx < len(values) else None
    except Exception:
        return None


def _make_collection(name, objects):
    col = Base()
    col["speckle_type"] = "Speckle.Core.Models.Collections.Collection"
    col["name"]         = name
    col["elements"]     = objects
    return col


def transfer_analysis_model(
    automate_context,
    speckle_client,
    version_root,
    target_stream_id,
    target_branch="main",
    isr_month="January",
):
    print(f"[Analysis Transfer] Starting → model: {target_stream_id}, month: {isr_month}")

    facade_objects = get_collection_objects(version_root, "Facade")
    slab_objects   = get_collection_objects(version_root, "Slabs")
    print(f"[Analysis Transfer] Facade:{len(facade_objects)} Slabs:{len(slab_objects)}")

    # 1. Extract ISR values
    isr_values = []
    for obj in facade_objects:
        v = _get_isr_value(obj, isr_month)
        if v is not None:
            isr_values.append(v)

    if not isr_values:
        print("[Analysis Transfer] WARNING: No ISR values found.")
        min_val = max_val = 0.0
    else:
        min_val = min(isr_values)
        max_val = max(isr_values)
        print(f"[Analysis Transfer] ISR range ({isr_month}): {min_val} → {max_val}")

    # 2. Apply colour per facade mesh
    val_range = max_val - min_val if max_val != min_val else 1.0
    for obj in facade_objects:
        v = _get_isr_value(obj, isr_month)
        t = (v - min_val) / val_range if v is not None else 0.0
        colour = _lerp_colour(t)
        obj["renderMaterial"] = RenderMaterial(
            name=f"ISR_{isr_month}_{round(v or 0, 1)}",
            diffuse=colour,
            opacity=1.0,
        )

    print(f"[Analysis Transfer] Colours applied.")

    # 3. Build root with Facade + Slabs
    new_root = Base()
    new_root["speckle_type"] = "Speckle.Core.Models.Collections.Collection"
    new_root["name"]         = "Analysis Model"
    new_root["elements"]     = [
        _make_collection("Facade", facade_objects),
        _make_collection("Slabs",  slab_objects),
    ]

    project_id = automate_context.automation_run_data.project_id
    print(f"[Analysis Transfer] Sending...")

    try:
        transport = ServerTransport(stream_id=project_id, client=speckle_client)
        obj_id = operations.send(base=new_root, transports=[transport])
        print(f"[Analysis Transfer] Sent: {obj_id}")
    except Exception as e:
        import traceback
        print(f"[Analysis Transfer] SEND ERROR: {type(e).__name__}: {e}")
        traceback.print_exc()
        raise

    try:
        new_version_id = automate_context.create_new_version_in_project(
            root_object=new_root,
            model_id=target_stream_id,
            version_message=f"Automate: facade ISR heatmap — {isr_month}",
        )
        print(f"[Analysis Transfer] Version created: {new_version_id}")
        return new_version_id
    except Exception as e:
        import traceback
        print(f"[Analysis Transfer] VERSION ERROR: {type(e).__name__}: {e}")
        traceback.print_exc()
        raise