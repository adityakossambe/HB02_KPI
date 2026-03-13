"""
transfer_modularity_model.py
-----------------------------
Transfers all 5 collections from the source (core) model to the
modularity target model.

Facade and Exoskeleton meshes get a colour assigned to properties.colors
based on the normalised modularity value of their ID:
  - Facade:      blue (most repeated) → orange (least repeated)
  - Exoskeleton: green (most repeated) → red (least repeated)

Slabs, Columns, Cores are transferred as-is.

The normalised value per ID is:
    unit_count_for_id / total_elements  (Facade + Exoskeleton combined)
"""

import copy
from collections import defaultdict

from specklepy.objects.base import Base
from specklepy.transports.server import ServerTransport
from specklepy.api import operations

from _collection_helper import get_collection_objects, get_prop, id_sort_key


# ── Colour helpers ────────────────────────────────────────────────────────

def _lerp(a, b, t):
    """Linear interpolate between a and b by factor t (0.0 → 1.0)."""
    return int(a + (b - a) * t)


def _to_argb(r, g, b, a=255):
    """Pack ARGB as a single integer (Speckle standard)."""
    return (a << 24) | (r << 16) | (g << 8) | b


# Colour endpoints for each scale
# Facade:      blue (high) → orange (low)
FACADE_HIGH = (0,   120, 255)   # blue
FACADE_LOW  = (255, 140,   0)   # orange

# Exoskeleton: green (high) → red (low)
EXO_HIGH    = (0,   200,  80)   # green
EXO_LOW     = (220,  30,  30)   # red


def _colour_for_normalised(normalised: float, high_rgb, low_rgb) -> int:
    """
    Map a normalised value (0.0–1.0) to an ARGB integer.
    normalised=1.0 → high_rgb (most repeated)
    normalised=0.0 → low_rgb  (least repeated)
    """
    t = max(0.0, min(1.0, normalised))
    r = _lerp(low_rgb[0], high_rgb[0], t)
    g = _lerp(low_rgb[1], high_rgb[1], t)
    b = _lerp(low_rgb[2], high_rgb[2], t)
    return _to_argb(r, g, b)


# ── Normalised value calculation ──────────────────────────────────────────

def _compute_normalised_values(facade_objects, exo_objects):
    """
    Compute normalised value per ID across both collections.
    Returns two dicts: {id: normalised_value} for facade and exo.
    """
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

    total_elements = len(facade_objects) + len(exo_objects)
    if total_elements == 0:
        return {}, {}

    facade_normalised = {pid: count / total_elements for pid, count in facade_counts.items()}
    exo_normalised    = {mid: count / total_elements for mid, count in exo_counts.items()}

    return facade_normalised, exo_normalised


# ── Colour application ────────────────────────────────────────────────────

def _apply_colour(obj, colour_int: int):
    """Write colour_int into obj.properties.colors."""
    props = getattr(obj, "properties", None)
    if props is None:
        obj["properties"] = Base()
        props = obj["properties"]
    try:
        props["colors"] = colour_int
    except Exception:
        obj["colors"] = colour_int
    return obj


def _colour_facade_objects(facade_objects, facade_normalised):
    coloured = []
    for obj in facade_objects:
        pid = str(get_prop(obj, "panel_id") or "").strip()
        norm_val = facade_normalised.get(pid, 0.0)
        colour   = _colour_for_normalised(norm_val, FACADE_HIGH, FACADE_LOW)
        coloured.append(_apply_colour(obj, colour))
    return coloured


def _colour_exo_objects(exo_objects, exo_normalised):
    coloured = []
    for obj in exo_objects:
        mid = str(get_prop(obj, "member_id") or "").strip()
        norm_val = exo_normalised.get(mid, 0.0)
        colour   = _colour_for_normalised(norm_val, EXO_HIGH, EXO_LOW)
        coloured.append(_apply_colour(obj, colour))
    return coloured


# ── Collection builder ────────────────────────────────────────────────────

def _make_collection(name: str, objects: list) -> Base:
    """Wrap a list of objects into a named Speckle collection."""
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
    """
    Transfer all 5 collections to the modularity target model.
    Facade and Exoskeleton get colour overrides on properties.colors.

    Parameters
    ----------
    automate_context  : AutomationContext from main.py
    speckle_client    : speckle_client from automate_context
    version_root      : received version root object
    target_stream_id  : stream ID of the modularity target model
    target_branch     : branch to commit to (default 'main')
    """

    print(f"[Modularity Transfer] Starting → stream: {target_stream_id}")

    # ── 1. Collect all 5 collections ─────────────────────────────────────
    slab_objects   = get_collection_objects(version_root, "Slabs")
    column_objects = get_collection_objects(version_root, "Columns")
    core_objects   = get_collection_objects(version_root, "Cores")
    facade_objects = get_collection_objects(version_root, "Facade")
    exo_objects    = get_collection_objects(version_root, "Exoskeleton")

    print(f"[Modularity Transfer] Objects — Slabs:{len(slab_objects)} Columns:{len(column_objects)} "
          f"Cores:{len(core_objects)} Facade:{len(facade_objects)} Exo:{len(exo_objects)}")

    # ── 2. Compute normalised values ──────────────────────────────────────
    facade_normalised, exo_normalised = _compute_normalised_values(
        facade_objects, exo_objects
    )

    # ── 3. Apply colours to Facade and Exoskeleton ────────────────────────
    coloured_facade = _colour_facade_objects(facade_objects, facade_normalised)
    coloured_exo    = _colour_exo_objects(exo_objects, exo_normalised)

    # ── 4. Build new root collection ──────────────────────────────────────
    new_root = Base()
    new_root["speckle_type"] = "Speckle.Core.Models.Collections.Collection"
    new_root["name"]     = "Grasshopper Model"
    new_root["elements"] = [
        _make_collection("Slabs",       slab_objects),
        _make_collection("Columns",     column_objects),
        _make_collection("Cores",       core_objects),
        _make_collection("Facade",      coloured_facade),
        _make_collection("Exoskeleton", coloured_exo),
    ]

    # ── 5. Send to target stream ──────────────────────────────────────────
    # project_id is the stream ID for transport and commit.
    # target_stream_id (model ID) is used as the branch name.
    project_id = automate_context.automation_run_data.project_id

    transport = ServerTransport(
        stream_id=project_id,
        client=speckle_client,
    )

    try:
        obj_id = operations.send(base=new_root, transports=[transport])
        print(f"[Modularity Transfer] Sent object: {obj_id}")
    except Exception as e:
        print(f"[Modularity Transfer] SEND FAILED: {type(e).__name__}: {e}")
        raise

    try:
        commit_id = speckle_client.commit.create(
            stream_id=project_id,
            object_id=obj_id,
            branch_name=target_branch,
            message="Automate: modularity model — colour coded by repetition index",
            source_application="SpeckleAutomate",
        )
        print(f"[Modularity Transfer] Committed: {commit_id}")
        return commit_id
    except Exception as e:
        print(f"[Modularity Transfer] COMMIT FAILED: {type(e).__name__}: {e}")
        raise