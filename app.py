"""
Audit Scheduling Assistant — Streamlit UI (v2, local-only)

Run: streamlit run app.py
Data: all inputs, outputs, cache, and run history stay on the local filesystem.
"""

from __future__ import annotations

import io
import sys
from datetime import date, timedelta
from pathlib import Path

import pandas as pd
import streamlit as st

sys.path.insert(0, str(Path(__file__).parent))
import audit_scheduler as sched  # noqa: E402
import history  # noqa: E402

st.set_page_config(page_title="Audit Scheduling Assistant", page_icon="📋", layout="wide")

# ----------------------------------------------------------------------------
# Sidebar — config + inputs
# ----------------------------------------------------------------------------

st.sidebar.title("⚙️ Configuration")

with st.sidebar.expander("How members are ranked", expanded=False):
    st.caption(
        "Selection is by fixed priority, not tunable weights:\n\n"
        "1. **Domain fit** — primary champion of a matched domain, then secondary.\n"
        "2. **Availability** — fewer leave days in the window.\n"
        "3. **Interest** — nominated this audit.\n\n"
        "TL must be a SAM who is a primary champion of a matched domain "
        "(closest partial match if none fits)."
    )

with st.sidebar.expander("Team split", expanded=False):
    primary_share = st.slider("Primary share", 0.30, 0.80, 0.50, 0.05)
    second_share = st.slider("Second share", 0.15, 0.50, 0.30, 0.05)
    team_floor = st.slider("Min floor per team", 0.05, 0.25, 0.15, 0.05)
    st.caption("Note: a 'Primary Audit Days' value in planned_audits.xlsx "
                "overrides the primary share for that audit. 'Analytics Used' "
                "is external and does not affect the split.")

with st.sidebar.expander("Window settings", expanded=False):
    lead_days = st.slider("Member lead before report (days)", 0, 150, 90, 15)
    days_per_week = st.slider("Audit-days per working week", 2, 5, 3, 1)

st.sidebar.divider()
st.sidebar.markdown("### 🗓️ Scheduling horizon")
quarter_choice = st.sidebar.selectbox(
    "Quarter to schedule",
    options=list(sched.QUARTERS.keys()),
    index=0,
    help="Only audits whose report date falls in this window are scheduled.",
)
_qstart, _qend = sched.QUARTERS[quarter_choice]
sched.set_horizon(_qstart, _qend)
st.sidebar.caption(f"Horizon: {_qstart:%d %b %Y} → {_qend:%d %b %Y}")

st.sidebar.divider()
st.sidebar.markdown("### 📁 Inputs")

data_source = st.sidebar.radio("Data source",
                                ["Use bundled test data", "Upload my own files"], index=0)

REQUIRED_FILES = ["leave_tracker", "domain_champions", "planned_audits",
                  "h1_allocations", "interests"]
OPTIONAL_FILES = ["skills"]
uploaded: dict[str, io.BytesIO] = {}
if data_source == "Upload my own files":
    for f in REQUIRED_FILES:
        u = st.sidebar.file_uploader(f"{f}.xlsx", type=["xlsx"], key=f)
        if u:
            uploaded[f] = u
    for f in OPTIONAL_FILES:
        u = st.sidebar.file_uploader(f"{f}.xlsx (optional)", type=["xlsx"], key=f)
        if u:
            uploaded[f] = u

st.sidebar.divider()
st.sidebar.markdown("### 📥 Import existing tracker")
st.sidebar.caption(
    "Upload an APS tracker (same format as `assignments_grid.xlsx`) to load "
    "existing assignments directly — skip the solver and jump straight to editing."
)
tracker_upload = st.sidebar.file_uploader(
    "APS tracker (.xlsx)", type=["xlsx"], key="aps_tracker_upload"
)

# ----------------------------------------------------------------------------
# Main pane
# ----------------------------------------------------------------------------

st.title("📋 Audit Scheduling Assistant")
st.caption("H2 2026 scheduling · respects H1 commitments · ICS / T&A / DM")

# Apply config to scheduler module
sched.BASE_SPLIT = {"primary": primary_share, "second": second_share,
                    "third": max(0.05, 1.0 - primary_share - second_share)}
sched.TEAM_FLOOR = team_floor
sched.MEMBER_LEAD_DAYS = lead_days
sched.AUDIT_DAYS_PER_WEEK = days_per_week


def _stage_inputs() -> Path:
    if data_source == "Use bundled test data":
        target = Path("./streamlit_inputs_bundled")
        marker = target / "domain_champions.xlsx"
        if not marker.exists():
            target.mkdir(exist_ok=True)
            # Use the Q1 2027 builder so bundled data matches the default horizon.
            import build_q1_2027 as gen
            old_out = gen.OUT
            gen.OUT = target
            try:
                gen.write_leave_tracker()
                gen.write_domain_champions()
                gen.write_planned_audits()
                gen.write_skills()
                gen.write_interests()
            finally:
                gen.OUT = old_out
        return target

    # Upload path
    target = Path("./streamlit_inputs_uploaded")
    target.mkdir(exist_ok=True)
    for fname, buf in uploaded.items():
        (target / f"{fname}.xlsx").write_bytes(buf.getvalue())
    return target


# Readiness: bundled = always OK; upload = REQUIRED files must all be present
REQUIRED = {"leave_tracker", "domain_champions", "planned_audits", "interests"}
OPTIONAL = {"h1_allocations"}
missing_required = REQUIRED - set(uploaded.keys())

if data_source == "Use bundled test data":
    ready = True
else:
    ready = not missing_required
    with st.sidebar:
        st.markdown("**Upload checklist**")
        for f in sorted(REQUIRED | OPTIONAL):
            tag = "(optional)" if f in OPTIONAL else "(required)"
            icon = "✅" if f in uploaded else ("⚪️" if f in OPTIONAL else "❌")
            st.caption(f"{icon} {f}.xlsx {tag}")

col_run, col_import, col_status = st.columns([1, 1, 3])
with col_run:
    run = st.button("▶️ Run scheduler", type="primary", disabled=not ready,
                     use_container_width=True)
with col_import:
    import_tracker = st.button(
        "📥 Load tracker", type="secondary",
        disabled=(tracker_upload is None),
        use_container_width=True,
        help="Load assignments from the uploaded APS tracker file.",
    )
with col_status:
    if not ready and tracker_upload is None:
        missing_list = ", ".join(sorted(missing_required))
        st.warning(f"Missing required files: **{missing_list}**. "
                    "Upload them, switch to bundled test data, or load a tracker.")
    elif not ready and tracker_upload is not None:
        st.info("A tracker is ready to import — click **Load tracker**.")

if "results" not in st.session_state:
    st.session_state.results = None


def _run():
    input_dir = _stage_inputs()
    sched.INPUT_DIR = input_dir
    out_dir = Path("./streamlit_outputs")
    out_dir.mkdir(exist_ok=True)
    sched.OUTPUT_DIR = out_dir
    cache_dir = Path("./streamlit_cache")
    cache_dir.mkdir(exist_ok=True)
    sched.CACHE_DIR = cache_dir
    sched.SKILL_CACHE_FILE = cache_dir / "skill_inference.json"

    with st.status("Running scheduler...", expanded=True) as status:
        st.write("Loading auditors (roster + leaves from leave tracker, "
                 "domains from champions file, H1 bookings, interests)...")
        auditors = sched.load_auditors()
        st.write(f"  → {len(auditors)} auditors")

        st.write("Loading audits in horizon...")
        audits = sched.load_audits()
        st.write(f"  → {len(audits)} audits in horizon")

        st.write("Computing durations and team-day splits...")
        for a in audits:
            sched.compute_duration(a)
            sched.compute_split(a)

        st.write("Matching audit titles to domains...")
        sched.infer_skills_for_audits(audits)

        st.write("Validating inputs...")
        findings = sched.validate_inputs(auditors, audits)
        n_err = sum(1 for f in findings if f["level"] == "error")
        n_warn = sum(1 for f in findings if f["level"] == "warning")
        st.write(f"  → {n_err} errors, {n_warn} warnings")

        st.write("Solving assignments...")
        warnings = sched.solve(auditors, audits)
        st.write(f"  → {len(warnings)} scheduling warnings")

        st.write("Writing output workbooks...")
        sched.write_grid(auditors, audits, out_dir / "assignments_grid.xlsx")
        sched.write_summary(audits, auditors, out_dir / "assignments_summary.xlsx")
        sched.write_warnings(warnings, auditors, out_dir / "warnings.xlsx")

        st.write("Archiving run to SQLite history...")
        run_id = history.new_run_id()
        config = {
            "split": {"primary": primary_share, "second": second_share,
                       "team_floor": team_floor},
            "window": {"lead_days": lead_days, "days_per_week": days_per_week},
            "horizon": quarter_choice,
            "data_source": data_source,
        }
        history.save_run(run_id, auditors, audits, warnings, config)
        status.update(label="✅ Done", state="complete", expanded=False)

    return {"auditors": auditors, "audits": audits, "warnings": warnings,
            "findings": findings, "out_dir": out_dir, "run_id": run_id}


def _load_tracker():
    """Load the uploaded APS tracker as the source of truth for assignments.

    Strategy:
      - If the scheduler input files are available (bundled or uploaded), load
        the full auditor roster + audit definitions from them first, so the
        editing tools have complete context (domains, leave, etc.).
      - Then overlay the tracker's bookings on top, replacing any solver output.
      - If input files aren't available, build a minimal context from the tracker
        itself (only name/designation/team from the tracker columns, no leave data).
    """
    out_dir = Path("./streamlit_outputs")
    out_dir.mkdir(exist_ok=True)

    with st.status("Loading tracker...", expanded=True) as status:
        # Best effort: load input context if available
        if ready:
            st.write("Loading auditor roster and audit definitions...")
            input_dir = _stage_inputs()
            sched.INPUT_DIR = input_dir
            cache_dir = Path("./streamlit_cache"); cache_dir.mkdir(exist_ok=True)
            sched.CACHE_DIR = cache_dir
            sched.SKILL_CACHE_FILE = cache_dir / "skill_inference.json"
            auditors = sched.load_auditors()
            audits = sched.load_audits()
            for a in audits:
                sched.compute_duration(a); sched.compute_split(a)
            sched.infer_skills_for_audits(audits)
            st.write(f"  → {len(auditors)} auditors, {len(audits)} audits")
        else:
            st.write("Input files not provided — building minimal context from tracker.")
            auditors = {}
            audits = []

        st.write("Parsing uploaded tracker...")
        tracker_upload.seek(0)
        auditors, audits, import_warnings = sched.load_tracker_upload(
            tracker_upload, auditors, audits
        )
        st.write(f"  → {len(auditors)} auditors after import, "
                 f"{len(import_warnings)} import notes")

        st.write("Writing output workbooks...")
        sched.OUTPUT_DIR = out_dir
        sched.write_grid(auditors, audits, out_dir / "assignments_grid.xlsx")
        sched.write_summary(audits, auditors, out_dir / "assignments_summary.xlsx")
        sched.write_warnings(import_warnings, auditors, out_dir / "warnings.xlsx")

        st.write("Archiving import to SQLite history...")
        run_id = history.new_run_id()
        history.save_run(run_id, auditors, audits, import_warnings,
                         {"source": "tracker_upload"})
        status.update(label="✅ Tracker loaded", state="complete", expanded=False)

    return {"auditors": auditors, "audits": audits, "warnings": import_warnings,
            "findings": [], "out_dir": out_dir, "run_id": run_id}


if run:
    st.session_state.results = _run()

if import_tracker and tracker_upload is not None:
    st.session_state.results = _load_tracker()

# ----------------------------------------------------------------------------
# Results
# ----------------------------------------------------------------------------

if st.session_state.results is None:
    st.markdown("---")
    st.markdown("""
    ### How it works

    **Option A — Run the scheduler**
    1. Roster comes from `leave_tracker.xlsx` (role, team, daily leaves).
    2. Domain ownership comes from `domain_champions.xlsx`.
    3. H2 audits from `planned_audits.xlsx`; H1 bookings from `h1_allocations.xlsx`; preferences from `interests.xlsx`.
    4. Each audit title is matched to domains by word overlap; domain champions form the eligible pool.
    5. TL = primary domain-champion SAM. Members ranked by domain fit → availability → interest.
    6. Members may start up to 3 months before the report date. No double-booking.

    **Option B — Load your existing tracker**
    Upload an APS tracker (sidebar → *Import existing tracker*) with columns:
    `Auditor Name | Designation | Team | 03 Aug 2026 | 10 Aug 2026 | ...`
    Cell values are audit titles (semicolons for multiple) or a leave code.
    Click **Load tracker** to import — all editing tools work the same way.

    Click **▶️ Run scheduler** or **📥 Load tracker** to begin.
    """)
    st.stop()

results = st.session_state.results
auditors = results["auditors"]
audits = results["audits"]
warnings_list = results["warnings"]
out_dir: Path = results["out_dir"]

# Source badge
_config = history.get_run_config(results["run_id"])
_src = _config.get("source", "scheduler")
if _src == "tracker_upload":
    st.info("📥 **Viewing imported tracker** — assignments loaded from your uploaded file. "
             "Use the editing tools below to make changes.")
else:
    st.success("🤖 **Viewing scheduler output** — assignments computed by the solver. "
                "Use the editing tools to override.")

# ---------- Top metrics ----------
m1, m2, m3, m4, m5 = st.columns(5)
m1.metric("Audits", len(audits))
m2.metric("Auditors", len(auditors))
m3.metric("Audits with TL",
          f"{sum(1 for a in audits if a.assigned_tl)}/{len(audits)}")
pref_total = sum(1 for a in auditors.values() if a.preferences)
pref_met = sum(1 for a in auditors.values() if a.preference_satisfied)
m4.metric("Preferences met", f"{pref_met}/{pref_total}")
m5.metric("Warnings", len(warnings_list),
          delta_color="inverse" if warnings_list else "off")

# ---------- Validation panel ----------
findings = results.get("findings", [])
errors = [f for f in findings if f["level"] == "error"]
warns_v = [f for f in findings if f["level"] == "warning"]
if findings:
    with st.expander(
        f"🔎 Input validation — {len(errors)} error(s), {len(warns_v)} warning(s)",
        expanded=bool(errors),
    ):
        if errors:
            st.markdown("**Errors** (these produce wrong or blank output — fix first):")
            for f in errors:
                aud = f"`{f['audit']}` " if f["audit"] else ""
                st.error(f"{aud}{f['message']}")
        if warns_v:
            st.markdown("**Warnings** (review recommended):")
            for f in warns_v:
                aud = f"`{f['audit']}` " if f["audit"] else ""
                st.warning(f"{aud}{f['message']}")
        if not errors and not warns_v:
            st.success("No input issues detected.")
else:
    st.success("🔎 Input validation passed — no issues detected.")

st.markdown("---")

tab_summary, tab_grid, tab_workload, tab_warnings, tab_downloads, tab_history = st.tabs([
    "📋 Per-Audit Summary",
    "📅 Calendar Grid",
    "👥 Auditor Workload",
    "⚠️ Warnings",
    "⬇️ Downloads",
    "📚 History",
])

# ---------- Per-Audit Summary ----------
with tab_summary:
    st.subheader("H2 audits")
    rows = []
    for a in sorted(audits, key=lambda x: x.report_date):
        scores = []
        if a.assigned_tl and a.assigned_tl in auditors:
            scores.append(sched.score(auditors[a.assigned_tl], a)[1]["skill"])
        for m_name, _, _, _ in a.assigned_members:
            if m_name in auditors:
                scores.append(sched.score(auditors[m_name], a)[1]["skill"])
        avg = round(sum(scores)/len(scores), 2) if scores else 0.0
        rows.append({
            "Audit": a.number,
            "Title": a.title,
            "Primary": a.primary_team,
            "Report": a.report_date.isoformat(),
            "Window": f"{a.start_date.strftime('%d %b')} → {a.end_date.strftime('%d %b')}",
            "Days": a.total_days,
            "Split (ICS/T&A/DM)": f"{a.team_days.get('ICS',0)}/{a.team_days.get('T&A',0)}/{a.team_days.get('DM',0)}",
            "TL": a.assigned_tl or "—",
            "Members": ", ".join(n for n, _, _, _ in a.assigned_members),
            "Skill Score": avg,
        })
    df = pd.DataFrame(rows)
    st.dataframe(df, use_container_width=True, hide_index=True,
                 column_config={"Skill Score": st.column_config.ProgressColumn(
                     "Skill Score", min_value=0, max_value=1, format="%.2f")})

    st.markdown("#### Drill into an audit")
    selected = st.selectbox(
        "Select an audit",
        options=[a.number for a in audits],
        format_func=lambda n: f"{n} — {next(a.title for a in audits if a.number == n)}",
    )
    audit = next(a for a in audits if a.number == selected)
    c1, c2 = st.columns(2)
    with c1:
        st.markdown(f"**{audit.title}**")
        st.markdown(f"Primary: `{audit.primary_team}` · Report: `{audit.report_date}` · "
                    f"Quarter: `{audit.quarter}`")
        st.markdown(f"Total days: `{audit.total_days}` · Calendar window: "
                    f"`{audit.duration_calendar_days}` days")
        st.markdown(f"Window: `{audit.start_date}` → `{audit.end_date}`")
        st.markdown(f"Split — ICS=`{audit.team_days.get('ICS',0)}`, "
                    f"T&A=`{audit.team_days.get('T&A',0)}`, "
                    f"DM=`{audit.team_days.get('DM',0)}`")
        st.markdown(f"Headcount — ICS=`{audit.team_headcount.get('ICS',0)}`, "
                    f"T&A=`{audit.team_headcount.get('T&A',0)}`, "
                    f"DM=`{audit.team_headcount.get('DM',0)}`")
        if audit.matched_domains:
            st.markdown("Matched domains: " +
                        " ".join(f"`{d}`" for d in audit.matched_domains[:4]))
    with c2:
        st.markdown("**Assignments**")
        rows = []
        if audit.assigned_tl:
            aud = auditors[audit.assigned_tl]
            _, b = sched.score(aud, audit)
            rows.append({"Role": "TL", "Name": aud.name, "Team": aud.home_team,
                          "Window": f"{audit.start_date.strftime('%d %b')} → {audit.end_date.strftime('%d %b')}",
                          "Skill": round(b["skill"], 2), "Prior": round(b["prior"], 2),
                          "Avail": round(b["avail"], 2), "Int": int(b["interest"])})
        for m_name, m_team, s, e in audit.assigned_members:
            if m_name in auditors:
                aud = auditors[m_name]
                _, b = sched.score(aud, audit, window=(s, e))
                rows.append({"Role": aud.role, "Name": aud.name, "Team": aud.home_team,
                              "Window": f"{s.strftime('%d %b')} → {e.strftime('%d %b')}",
                              "Skill": round(b["skill"], 2), "Prior": round(b["prior"], 2),
                              "Avail": round(b["avail"], 2), "Int": int(b["interest"])})
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

    # ----- Manual overrides -----
    st.markdown("---")
    st.markdown("#### ✏️ Manual overrides")
    st.caption(
        "Surgical edits to this audit's assignment. Validations are advisory — "
        "you can override warnings if you know something the scheduler doesn't."
    )

    # Helper: rewrite an auditor's H2 booking for this audit
    def _rewrite_booking(person: sched.Auditor, new_start: date, new_end: date):
        person.bookings = [b for b in person.bookings
                            if not (b.source == "H2" and b.audit_number == audit.number)]
        person.bookings.append(sched.Booking(
            audit_number=audit.number, audit_title=audit.title,
            start=new_start, end=new_end, source="H2",
        ))

    def _drop_booking(person: sched.Auditor):
        person.bookings = [b for b in person.bookings
                            if not (b.source == "H2" and b.audit_number == audit.number)]

    def _conflict_note(person: sched.Auditor, start: date, end: date) -> str | None:
        """Return a human-readable conflict warning, or None if clear."""
        conflicts = []
        for b in person.bookings:
            if b.source == "H2" and b.audit_number == audit.number:
                continue  # the slot we're about to overwrite
            if not (b.end < start or b.start > end):
                tag = "H1" if b.source == "H1" else "H2"
                conflicts.append(f"{tag}:{b.audit_title} ({b.start}→{b.end})")
        leave_overlap = sum(1 for d in person.leaves if start <= d <= end and d.weekday() < 5)
        msg_parts = []
        if conflicts:
            msg_parts.append("overlaps: " + "; ".join(conflicts))
        if leave_overlap:
            msg_parts.append(f"{leave_overlap} leave-day(s) in window")
        return " · ".join(msg_parts) if msg_parts else None

    oc1, oc2 = st.columns(2)

    # --- Swap TL ---
    with oc1:
        st.markdown("**Swap Team Lead**")
        sam_options = [a.name for a in auditors.values()
                        if a.is_sam and a.home_team == audit.primary_team]
        current_tl = audit.assigned_tl or "(unassigned)"
        new_tl = st.selectbox(
            f"Team Lead (SAM in {audit.primary_team})",
            options=["(unassigned)"] + sam_options,
            index=(sam_options.index(audit.assigned_tl) + 1)
                   if audit.assigned_tl in sam_options else 0,
            key=f"tl_{audit.number}",
        )
        if new_tl != current_tl:
            note = (_conflict_note(auditors[new_tl], audit.start_date, audit.end_date)
                    if new_tl != "(unassigned)" else None)
            if note:
                st.warning(f"⚠️ {new_tl}: {note}")
            if st.button("Apply TL change", key=f"apply_tl_{audit.number}"):
                old_tl = audit.assigned_tl
                if audit.assigned_tl and audit.assigned_tl in auditors:
                    _drop_booking(auditors[audit.assigned_tl])
                if new_tl != "(unassigned)":
                    _rewrite_booking(auditors[new_tl], audit.start_date, audit.end_date)
                    audit.assigned_tl = new_tl
                else:
                    audit.assigned_tl = None
                history.log_override(
                    results["run_id"], "swap_tl", audit.number,
                    {"from": old_tl, "to": audit.assigned_tl},
                )
                st.success("TL updated.")
                st.rerun()

    # --- Manage members ---
    with oc2:
        st.markdown("**Manage members**")
        current_members = {n: (s, e) for n, _, s, e in audit.assigned_members}

        # Remove
        if current_members:
            to_remove = st.multiselect(
                "Remove from team",
                options=list(current_members.keys()),
                key=f"rm_{audit.number}",
            )
            if to_remove and st.button("Remove selected", key=f"do_rm_{audit.number}"):
                for name in to_remove:
                    if name in auditors:
                        _drop_booking(auditors[name])
                audit.assigned_members = [m for m in audit.assigned_members
                                            if m[0] not in to_remove]
                history.log_override(
                    results["run_id"], "remove_member", audit.number,
                    {"removed": to_remove},
                )
                st.success(f"Removed: {', '.join(to_remove)}")
                st.rerun()

        # Add
        st.markdown("Add member")
        eligible = [a for a in auditors.values()
                     if a.name not in current_members and a.name != audit.assigned_tl]
        eligible_sorted = sorted(eligible, key=lambda x: (x.home_team, x.role, x.name))
        new_member = st.selectbox(
            "Auditor to add",
            options=["(select)"] + [f"{a.name} — {a.home_team}/{a.role}" for a in eligible_sorted],
            key=f"add_{audit.number}",
        )
        ac1, ac2 = st.columns(2)
        _lead_floor = max(sched.H2_START,
                          audit.report_date - timedelta(days=sched.MEMBER_LEAD_DAYS))
        with ac1:
            add_start = st.date_input("Start",
                value=max(_lead_floor, audit.start_date),
                min_value=_lead_floor,
                max_value=audit.end_date,
                key=f"add_start_{audit.number}",
            )
        with ac2:
            add_end = st.date_input("End",
                value=audit.end_date,
                min_value=_lead_floor,
                max_value=audit.end_date,
                key=f"add_end_{audit.number}",
            )
        if new_member != "(select)":
            picked_name = new_member.split(" — ")[0]
            note = _conflict_note(auditors[picked_name], add_start, add_end)
            if note:
                st.warning(f"⚠️ {picked_name}: {note}")
            if st.button("Add to team", key=f"do_add_{audit.number}"):
                if add_end < add_start:
                    st.error("End date must be on or after start date.")
                else:
                    person = auditors[picked_name]
                    _rewrite_booking(person, add_start, add_end)
                    audit.assigned_members.append((person.name, person.home_team, add_start, add_end))
                    history.log_override(
                        results["run_id"], "add_member", audit.number,
                        {"name": picked_name, "start": str(add_start), "end": str(add_end)},
                    )
                    st.success(f"Added {picked_name}.")
                    st.rerun()

    # --- Adjust existing member windows ---
    if audit.assigned_members:
        st.markdown("**Adjust member windows**")
        for idx, (m_name, m_team, m_start, m_end) in enumerate(list(audit.assigned_members)):
            wc1, wc2, wc3, wc4 = st.columns([2, 2, 2, 1])
            _lf = max(sched.H2_START,
                      audit.report_date - timedelta(days=sched.MEMBER_LEAD_DAYS))
            with wc1:
                st.markdown(f"`{m_name}` ({m_team})")
            with wc2:
                new_s = st.date_input(
                    "start", value=m_start, key=f"win_s_{audit.number}_{idx}",
                    min_value=min(_lf, m_start),
                    max_value=audit.end_date,
                    label_visibility="collapsed",
                )
            with wc3:
                new_e = st.date_input(
                    "end", value=m_end, key=f"win_e_{audit.number}_{idx}",
                    min_value=min(_lf, m_start),
                    max_value=max(audit.end_date, m_end),
                    label_visibility="collapsed",
                )
            with wc4:
                if st.button("Save", key=f"win_save_{audit.number}_{idx}"):
                    if new_e < new_s:
                        st.error("End must be ≥ start.")
                    else:
                        if m_name in auditors:
                            _rewrite_booking(auditors[m_name], new_s, new_e)
                        audit.assigned_members[idx] = (m_name, m_team, new_s, new_e)
                        history.log_override(
                            results["run_id"], "adjust_window", audit.number,
                            {"name": m_name,
                             "from": [str(m_start), str(m_end)],
                             "to": [str(new_s), str(new_e)]},
                        )
                        st.success(f"Updated {m_name}'s window.")
                        st.rerun()
            if m_name in auditors:
                note = _conflict_note(auditors[m_name], new_s, new_e)
                if note:
                    st.caption(f"⚠️ {note}")

    # --- Re-export with overrides ---
    st.markdown("**Persist changes**")
    rxc1, rxc2 = st.columns([1, 3])
    with rxc1:
        if st.button("🔄 Re-export Excel files", key=f"reexport_{audit.number}"):
            sched.write_grid(auditors, audits, out_dir / "assignments_grid.xlsx")
            sched.write_summary(audits, auditors, out_dir / "assignments_summary.xlsx")
            sched.write_warnings(warnings_list, auditors, out_dir / "warnings.xlsx")
            history.replace_assignments(results["run_id"], auditors, audits)
            st.success("Output files refreshed and SQLite history updated.")
    with rxc2:
        st.caption("Overrides apply immediately within the app. "
                    "Click Re-export to refresh the downloadable Excel files "
                    "and persist the overridden assignments in history.")

# ---------- Calendar Grid ----------
with tab_grid:
    st.subheader("Weekly grid (H1 grey, H2 green)")
    team = st.radio("Team", sched.TEAMS, horizontal=True, key="grid_team")
    first_monday = date(sched.SCHEDULE_YEAR, 1, 5)
    weeks = [first_monday + timedelta(weeks=i) for i in range(52)]

    team_auditors = sorted(
        [a for a in auditors.values() if a.home_team == team],
        key=lambda x: (x.role != "SAM", x.role, x.name),
    )
    grid = []
    for aud in team_auditors:
        row = {"Auditor": f"{aud.role}/{aud.name}"}
        for w in weeks:
            week_end = w + timedelta(days=4)
            titles = []
            for b in aud.bookings:
                if b.start <= week_end and b.end >= w:
                    src = "[H1]" if b.source == "H1" else ""
                    titles.append(f"{src}{b.audit_title[:18]}")
            row[w.strftime("%d %b")] = "; ".join(titles)
        grid.append(row)
    df = pd.DataFrame(grid)

    def highlight(val):
        if not isinstance(val, str) or not val.strip():
            return ""
        if "[H1]" in val:
            return "background-color: #e8e8e8; color: #555"
        return "background-color: #c6e5b3; color: #1a4a1a"
    styled = df.style.map(highlight, subset=df.columns[1:])
    st.dataframe(styled, use_container_width=True, height=420, hide_index=True)

    # ------------------------------------------------------------------
    # Free-range booking editor
    # ------------------------------------------------------------------
    st.markdown("---")
    st.markdown("### ✏️ Edit bookings (any auditor, any range)")
    st.caption(
        "Pick anyone and assign them to any audit between any two dates — "
        "no audit-window constraints. Overlaps, leaves, and out-of-H2 ranges "
        "are shown as advisory warnings but never block you."
    )

    ec1, ec2 = st.columns([1, 1])
    with ec1:
        edit_auditor = st.selectbox(
            "Auditor",
            options=sorted(auditors.keys(),
                            key=lambda n: (auditors[n].home_team,
                                            auditors[n].role != "SAM",
                                            n)),
            format_func=lambda n: f"{auditors[n].name} — {auditors[n].home_team}/{auditors[n].role}",
            key="edit_auditor_sel",
        )
    with ec2:
        edit_audit = st.selectbox(
            "Audit",
            options=[a.number for a in audits],
            format_func=lambda n: next(
                (f"{a.number} — {a.title}" for a in audits if a.number == n), n),
            key="edit_audit_sel",
        )

    person = auditors[edit_auditor]
    target_audit = next(a for a in audits if a.number == edit_audit)

    # Existing bookings for this auditor — listed with remove buttons
    st.markdown(f"**Existing bookings for `{person.name}`**")
    if person.bookings:
        for i, b in enumerate(list(person.bookings)):
            tag = "🔒 H1" if b.source == "H1" else "🟢 H2"
            bc1, bc2, bc3, bc4 = st.columns([0.7, 2.5, 2, 0.8])
            with bc1:
                st.markdown(tag)
            with bc2:
                st.markdown(f"`{b.audit_number}` — {b.audit_title[:40]}")
            with bc3:
                st.markdown(f"{b.start} → {b.end}")
            with bc4:
                if st.button("Remove", key=f"rm_book_{person.name}_{i}"):
                    # Drop from auditor
                    removed = person.bookings.pop(i)
                    # Also remove from any audit's assigned_members / TL
                    for a in audits:
                        if a.number == removed.audit_number:
                            if a.assigned_tl == person.name:
                                a.assigned_tl = None
                            a.assigned_members = [m for m in a.assigned_members
                                                    if m[0] != person.name]
                    history.log_override(
                        results["run_id"], "free_remove", removed.audit_number,
                        {"name": person.name,
                         "from": [str(removed.start), str(removed.end)],
                         "source": removed.source},
                    )
                    st.success(f"Removed {person.name} from {removed.audit_number}.")
                    st.rerun()
    else:
        st.caption("_(no bookings yet)_")

    # Add or replace a booking — totally free date range
    st.markdown(f"**Assign `{person.name}` to `{target_audit.number}`**")
    year_start = date(sched.SCHEDULE_YEAR, 1, 1)
    year_end = date(sched.SCHEDULE_YEAR, 12, 31)
    nc1, nc2, nc3 = st.columns([1, 1, 1])
    with nc1:
        new_start = st.date_input(
            "Start",
            value=target_audit.start_date,
            min_value=year_start, max_value=year_end,
            key=f"free_start_{person.name}_{target_audit.number}",
        )
    with nc2:
        new_end = st.date_input(
            "End",
            value=target_audit.end_date,
            min_value=year_start, max_value=year_end,
            key=f"free_end_{person.name}_{target_audit.number}",
        )
    with nc3:
        role_choice = st.selectbox(
            "Role on this audit",
            options=["Member", "Team Lead"],
            key=f"free_role_{person.name}_{target_audit.number}",
        )

    # Advisory warnings — never blocking
    advisory = []
    if new_end < new_start:
        advisory.append("End is before start.")
    if new_start < sched.H2_START or new_end > sched.H2_END:
        advisory.append("Range extends outside H2 (Jul–Dec 2026).")
    leave_days = sum(1 for d in person.leaves
                      if new_start <= d <= new_end and d.weekday() < 5)
    if leave_days:
        advisory.append(f"{leave_days} leave-day(s) in this window.")
    for b in person.bookings:
        if b.audit_number == target_audit.number:
            continue
        if not (b.end < new_start or b.start > new_end):
            tag = "H1" if b.source == "H1" else "H2"
            advisory.append(f"Overlaps existing {tag} booking: "
                             f"{b.audit_title} ({b.start}→{b.end})")
    if not person.is_sam and role_choice == "Team Lead":
        advisory.append(f"{person.name} is {person.role}, not a SAM — "
                         "unusual choice for Team Lead.")
    fit = sched.domain_fit(person, target_audit)
    if fit == 0 and target_audit.matched_domains:
        advisory.append("Not a domain champion (primary/secondary) for this audit.")
    for note in advisory:
        st.warning(f"⚠️ {note}")

    if st.button("✅ Apply assignment",
                  key=f"free_apply_{person.name}_{target_audit.number}",
                  type="primary",
                  disabled=(new_end < new_start)):
        # Remove any prior booking for this audit on this auditor
        person.bookings = [b for b in person.bookings
                            if b.audit_number != target_audit.number]
        # Drop from audit's assignment lists first
        if target_audit.assigned_tl == person.name:
            target_audit.assigned_tl = None
        target_audit.assigned_members = [m for m in target_audit.assigned_members
                                          if m[0] != person.name]
        # Add the new booking
        person.bookings.append(sched.Booking(
            audit_number=target_audit.number, audit_title=target_audit.title,
            start=new_start, end=new_end, source="H2",
        ))
        if role_choice == "Team Lead":
            # If there's already a TL, demote them to a member
            if target_audit.assigned_tl and target_audit.assigned_tl != person.name:
                old_tl = target_audit.assigned_tl
                if old_tl in auditors:
                    target_audit.assigned_members.append(
                        (old_tl, auditors[old_tl].home_team,
                         target_audit.start_date, target_audit.end_date))
            target_audit.assigned_tl = person.name
        else:
            target_audit.assigned_members.append(
                (person.name, person.home_team, new_start, new_end))
        history.log_override(
            results["run_id"], "free_assign", target_audit.number,
            {"name": person.name, "role": role_choice,
             "window": [str(new_start), str(new_end)]},
        )
        st.success(
            f"Assigned {person.name} to {target_audit.number} as {role_choice} "
            f"from {new_start} to {new_end}."
        )
        st.rerun()

# ---------- Workload ----------
with tab_workload:
    st.subheader("Auditor workload (H2 only)")
    rows = []
    for aud in sorted(auditors.values(), key=lambda x: (x.home_team, x.role, x.name)):
        rows.append({
            "Auditor": aud.name, "Team": aud.home_team, "Role": aud.role,
            "H2 audits": len(aud.assigned_h2_audits),
            "H2 days": aud.days_booked_h2,
            "Leave days (H2)": sum(1 for d in aud.leaves
                                    if sched.H2_START <= d <= sched.H2_END),
            "Pref met": ("✅" if aud.preference_satisfied
                          else ("—" if not aud.preferences else "❌")),
        })
    df_wl = pd.DataFrame(rows)
    c1, c2 = st.columns(2)
    with c1:
        st.dataframe(df_wl, use_container_width=True, hide_index=True,
                     column_config={"H2 days": st.column_config.ProgressColumn(
                         "H2 days", min_value=0, max_value=130, format="%d")})
    with c2:
        st.bar_chart(df_wl.set_index("Auditor")[["H2 days"]], height=420)

    st.markdown("#### Domain fit of assigned people")
    st.caption("Flags assignees who are neither primary nor secondary champion of "
                "any domain the audit matched. Some non-fit members are expected "
                "(an audit may legitimately borrow capacity), but a non-fit TL is worth review.")
    gap_rows = []
    for aud in auditors.values():
        rel = [a for a in audits
                if aud.name == a.assigned_tl
                or any(aud.name == n for n, _, _, _ in a.assigned_members)]
        for a in rel:
            fit = sched.domain_fit(aud, a)
            if fit == 0 and a.matched_domains:
                role = "TL" if aud.name == a.assigned_tl else "Member"
                gap_rows.append({"Auditor": aud.name, "Role": role,
                                  "Audit": a.number,
                                  "Audit title": a.title,
                                  "Matched domains": ", ".join(a.matched_domains[:3])})
    if gap_rows:
        df_gap = pd.DataFrame(gap_rows)
        # Surface non-fit TLs first
        df_gap = df_gap.sort_values(by="Role", ascending=True)
        st.dataframe(df_gap, use_container_width=True, hide_index=True)
    else:
        st.success("Every assignee is a domain champion (primary or secondary) of their audit.")

# ---------- Warnings ----------
with tab_warnings:
    st.subheader("Scheduler warnings")
    if not warnings_list:
        st.success("✅ No warnings.")
    else:
        for w in warnings_list:
            st.warning(w)

    st.markdown("#### H2 capacity")
    h2_audit_days = sum(a.total_days for a in audits)
    h2_work = sched.working_days_between(sched.H2_START, sched.H2_END)
    h1_booked = sum(sched.working_days_between(b.start, b.end)
                     for aud in auditors.values()
                     for b in aud.bookings if b.source == "H1")
    capacity_h2 = len(auditors) * h2_work
    util = h2_audit_days / max(1, capacity_h2)
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("H2 days planned", f"{h2_audit_days:,}")
    c2.metric("H1 days booked", f"{h1_booked:,}")
    c3.metric("H2 raw capacity", f"{capacity_h2:,}")
    c4.metric("H2 utilisation", f"{util:.0%}")
    st.progress(min(1.0, util))

# ---------- Downloads ----------
with tab_downloads:
    st.subheader("Download outputs")
    for fname, label in [
        ("assignments_summary.xlsx", "📋 Per-audit summary"),
        ("assignments_grid.xlsx", "📅 Weekly grid"),
        ("warnings.xlsx", "⚠️ Warnings + workload"),
    ]:
        p = out_dir / fname
        if p.exists():
            st.download_button(
                label=f"{label} ({fname})",
                data=p.read_bytes(),
                file_name=fname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

    # SQLite DB download
    db_path = history.DB_PATH
    if db_path.exists():
        st.markdown("---")
        st.markdown("**Run history database**")
        st.caption(f"All past runs are stored locally in `{db_path}`. "
                    "Download a snapshot of the SQLite file to share or archive.")
        st.download_button(
            label=f"📚 Run history ({db_path.name})",
            data=db_path.read_bytes(),
            file_name=db_path.name,
            mime="application/x-sqlite3",
            use_container_width=True,
        )

# ---------- History ----------
with tab_history:
    st.subheader("Past scheduler runs")
    st.caption(f"All runs are persisted locally in `{history.DB_PATH}`. "
                "Nothing here leaves your machine.")

    runs = history.list_runs(limit=100)
    if not runs:
        st.info("No past runs yet. Run the scheduler once and come back.")
    else:
        # Summary table
        st.dataframe(
            pd.DataFrame(runs).rename(columns={
                "run_id": "Run ID", "started_at": "Started",
                "n_auditors": "Auditors", "n_audits": "Audits",
                "n_warnings": "Warnings", "notes": "Notes",
            }),
            use_container_width=True, hide_index=True,
        )

        st.markdown("#### Drill into a past run")
        run_options = [r["run_id"] for r in runs]
        # Highlight the current run if it's in the list
        default_idx = 0
        if "run_id" in results and results["run_id"] in run_options:
            default_idx = run_options.index(results["run_id"])
        chosen_run = st.selectbox("Run", run_options, index=default_idx,
                                    format_func=lambda r: f"{r}  ({'current' if r == results.get('run_id') else 'past'})")
        cfg = history.get_run_config(chosen_run)
        if cfg:
            with st.expander("Config used"):
                st.json(cfg)

        assignments = history.get_run_assignments(chosen_run)
        warns_past = history.get_run_warnings(chosen_run)
        overrides = history.get_overrides_log(chosen_run)

        hc1, hc2 = st.columns(2)
        with hc1:
            st.markdown("**Assignments**")
            if assignments:
                df_a = pd.DataFrame(assignments)
                df_a["is_override"] = df_a["is_override"].map({0: "—", 1: "✏️"})
                st.dataframe(
                    df_a.rename(columns={"is_override": "Override"}),
                    use_container_width=True, hide_index=True, height=400,
                )
            else:
                st.info("No assignments recorded.")
        with hc2:
            st.markdown("**Warnings**")
            if warns_past:
                for w in warns_past:
                    st.warning(w)
            else:
                st.success("No warnings.")

            st.markdown("**Overrides log**")
            if overrides:
                for o in overrides:
                    st.caption(f"`{o['logged_at']}` · **{o['action']}** on "
                                f"`{o['audit_number']}` — {o['details']}")
            else:
                st.caption("_No manual overrides logged for this run._")

        st.markdown("---")
        del_col1, del_col2 = st.columns([1, 4])
        with del_col1:
            confirm_delete = st.checkbox("Confirm delete", key=f"confirm_del_{chosen_run}")
        with del_col2:
            if st.button("🗑️ Delete this run", disabled=not confirm_delete):
                history.delete_run(chosen_run)
                st.success(f"Deleted {chosen_run}.")
                st.rerun()
