"""
transfer_modularity_model.py
-----------------------------
Transfers all 5 collections to the modularity target model.

Facade and Exoskeleton mesh colours are set via the top-level `colors`
array (list of ARGB integers, one per vertex) based on normalised
modularity value of their ID:
  - Facade:      blue (most repeated) → orange (least repeated)
  - Exoskeleton: green (most repeated) → red (least repeated)

Slabs, Columns, Cores transferred as-is with no colour changes.
Objects modified in-place on the already-received version to avoid
resending all geometry data.
"""

from collections import defaultdict
from specklepy.objects.base import Base
from _collection_helper import get_collection_objects, get_prop


# ── Colour helpers ────────────────────────────────────────────────────────

def _lerp(a, b, t):
    return int(a + (b - a) * t)

def _to_argb(r, g, b, a=255):
    return (a << 24) | (r << 16) | (g << 8) | b

FACADE_HIGH = (0,   120, 255)   # blue
FACADE_LOW  = (255, 140,   0)   # orange
EXO_HIGH    = (0,   200,  80)   # green
EXO_LOW     = (220,  30,  30)   # red


def _colour_for_normalised(normalised: float, high_rgb, low_rgb) -> int:
    t = max(0.0, min(1.0, normalised))
    r = _lerp(low_rgb[0], high_rgb[0], t)
    g = _lerp(low_rgb[1], high_rgb[1], t)
    b = _lerp(low_rgb[2], high_rgb[2], t)
    return _to_argb(r, g, b)


def _apply_colour_to_mesh(obj, colour_int: int):
    """Fill the top-level colors array with one ARGB int per vertex."""
    try:
        vertices = getattr(obj, "vertices", None)
        vertex_count = len(vertices) // 3 if vertices and len(vertices) > 0 else 1
    except Exception:
        vertex_count = 1
    obj["colors"] = [colour_int] * vertex_count


# ── Normalised value calculation ──────────────────────────────────────────

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

def _make_collection(name: str, objects: list) -> Base:
    col = Base()
    col["speckle_type"] = "Speckle.Core.Models.Collections.Collection"
    col["name"]         = name
    col["elements"]     = objects
    return col


# ── Main transfer function ────────────────────────────────────────────────

def transfer_modularity_model(
    automate_context,
    speckle_client,
    version_root,
    target_stream_id: str,
    target_branch: str = "main",
):
    print(f"[Modularity Transfer] Starting → model: {target_stream_id}")

    # ── 1. Collect ───────────────────────────────────────────────────────
    slab_objects   = get_collection_objects(version_root, "Slabs")
    column_objects = get_collection_objects(version_root, "Columns")
    core_objects   = get_collection_objects(version_root, "Cores")
    facade_objects = get_collection_objects(version_root, "Facade")
    exo_objects    = get_collection_objects(version_root, "Exoskeleton")

    print(f"[Modularity Transfer] Slabs:{len(slab_objects)} Columns:{len(column_objects)} "
          f"Cores:{len(core_objects)} Facade:{len(facade_objects)} Exo:{len(exo_objects)}")

    # ── 2. Normalised values ─────────────────────────────────────────────
    facade_norm, exo_norm = _compute_normalised_values(facade_objects, exo_objects)

    # ── 3. Apply colours in-place ─────────────────────────────────────────
    for obj in facade_objects:
        pid  = str(get_prop(obj, "panel_id") or "").strip()
        colour = _colour_for_normalised(facade_norm.get(pid, 0.0), FACADE_HIGH, FACADE_LOW)
        _apply_colour_to_mesh(obj, colour)

    for obj in exo_objects:
        mid  = str(get_prop(obj, "member_id") or "").strip()
        colour = _colour_for_normalised(exo_norm.get(mid, 0.0), EXO_HIGH, EXO_LOW)
        _apply_colour_to_mesh(obj, colour)

    print("[Modularity Transfer] Colours applied, creating new version...")

    # ── 4. Build new root ────────────────────────────────────────────────
    new_root = Base()
    new_root["speckle_type"] = "Speckle.Core.Models.Collections.Collection"
    new_root["name"]         = "Grasshopper Model"
    new_root["elements"]     = [
        _make_collection("Slabs",       slab_objects),
        _make_collection("Columns",     column_objects),
        _make_collection("Cores",       core_objects),
        _make_collection("Facade",      facade_objects),
        _make_collection("Exoskeleton", exo_objects),
    ]

    # ── 5. Send via automate context ─────────────────────────────────────
    try:
        new_version_id = automate_context.create_new_version_in_project(
            root_object=new_root,
            model_id=target_stream_id,
            version_message="Automate: modularity model — colour coded by repetition index",
        )
        print(f"[Modularity Transfer] Done. Version: {new_version_id}")
        return new_version_id

    except Exception as e:
        print(f"[Modularity Transfer] FAILED: {type(e).__name__}: {e}")
        import traceback
        traceback.print_exc()
        raise