"""
Lightweight web UI for the SSAS -> Fabric migration tool.

Run with (see README Section 11 for full setup, including why this needs
its OWN isolated virtual environment):
    .venv-ui\\Scripts\\python.exe -m streamlit run ssas_fabric_migrator/ui/app.py

This is a thin wrapper around the existing CLI (`cli/orchestrator.py`,
`datamover/loader.py`) and its underlying modules - it does not duplicate
any pipeline logic. It just gives users a form-based way to configure
connections and click through each step instead of typing CLI commands,
with live log output, per-step explanations, and a built-in viewer for the
generated reports.

DEPLOYMENT NOTE:
This app must run on a Windows host with:
  - line-of-sight network access to the on-prem SSAS instance and SQL
    Server (same domain/VPN as required by the CLI today), and
  - an x64 Python interpreter with pythonnet/pandas/pyarrow/deltalake
    installed, configured via the "Python executable" field below (this is
    a SEPARATE interpreter from the one running Streamlit itself - see
    requirements-ui.txt for why they must not share an environment).
"Any device" access is achieved by hosting this single app once on such a
host and letting users reach it via browser (see README Section 11) - the
UI itself does not need to be installed on every analyst's machine.
"""
from __future__ import annotations

import html
import json
import os
import re
import subprocess
import sys

import streamlit as st

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
README_PATH = os.path.join(REPO_ROOT, "README.md")
# Used to rewrite README.md's in-file anchor links (e.g. "(#3-prerequisites)") into
# absolute GitHub links when rendering excerpts in the Read Me tab - GitHub auto-generates
# heading anchors so these resolve there, but Streamlit's markdown renderer does not, and
# the Read Me tab only shows excerpts (not the full file) so in-app anchors would be dead.
README_GITHUB_URL = (
    "https://github.com/sagarbathe/SSAS_MDXCube_to_Fabric_Migration_UI_Tool/blob/main/README.md"
)

# GitHub's auto-generated heading anchors for every README.md section this app refers users
# to by number/name - kept as a lookup instead of re-deriving them at runtime so every
# "see README Section N" mention in the UI can be a real clickable link instead of dead
# plain text. Must be kept in sync with README.md's actual "## N. Title" headings.
README_SECTION_ANCHORS = {
    "1": "1-purpose--objective",
    "2": "2-two-phase-design-on-prem-vs-fabric-connected",
    "3": "3-prerequisites",
    "4": "4-install-dependencies",
    "5": "5-phase-1-on-prem-ssas--sql-server",
    "6": "6-phase-2-fabric-connected",
    "7": "7-alternative-data-migration-options-for-review",
    "8": "8-the-migration-conversion-report",
    "9": "9-repository-structure",
    "10": "10-limitations",
    "11": "11-web-ui",
    "step-7": "step-7-deploy-the-semantic-model",
    "quickstart": "quickstart-test-the-tool-end-to-end",
}


def _readme_link(section_key: str, label: str) -> str:
    """Returns a Markdown link to a specific README.md section on GitHub, e.g.
    _readme_link("10", "Section 10") -> "[Section 10](https://.../README.md#10-limitations)".
    Used everywhere the UI mentions a README section by number so it's a real click-through
    instead of plain text the user has to go find manually."""
    anchor = README_SECTION_ANCHORS[section_key]
    return f"[{label}]({README_GITHUB_URL}#{anchor})"

ENV_FIELDS = [
    ("SSAS_SERVER", "On-prem SSAS server\\instance", False, "LAPTOP-LQVSA8VE\\SSAS"),
    ("SSAS_DATABASE", "On-prem SSAS database (cube) name", False, "RetailCubeDemo"),
    ("SQL_SERVER", "On-prem SQL Server (relational source)", False, "localhost"),
    ("SQL_DATABASE", "On-prem SQL Server database", False, "RetailDW"),
    ("FABRIC_TENANT_ID", "Fabric/Entra ID tenant ID", False, ""),
    ("FABRIC_CLIENT_ID", "Fabric app registration (service principal) client ID", False, ""),
    ("FABRIC_CLIENT_SECRET", "Fabric app registration client secret", True, ""),
    ("FABRIC_WORKSPACE_ID", "Target Fabric workspace ID", False, ""),
    ("SEMANTIC_MODEL_NAME", "Fabric Semantic Model name", False, "RetailCubeDemo"),
]

STEP_LABELS = {
    "extract": "1. Extract cube metadata",
    "analyze": "2. Analyze Direct Lake feasibility",
    "generate": "3. Generate TMDL + notebook scripts",
    "report": "4. Generate MIGRATION_REPORT.md",
    "deploy-lake": "5. Deploy/find Lakehouse",
    "migrate-data": "6. Migrate data (direct)",
    "upload-data": "6. Migrate data (offline transfer - step 2 of 2)",
    "deploy-model": "7. Deploy semantic model",
}

STEP_DESCRIPTIONS = {
    "extract": (
        "Connects to the on-prem SSAS Multidimensional cube via **AMO** (Windows "
        "Integrated Auth - the account running this app needs the SSAS **Server "
        "Administrator** role) and dumps its full metadata (dimensions, measures, "
        "hierarchies, calculated members, KPIs, DSV schema) to `cube_metadata.json`. "
        "No Fabric connectivity needed for this step."
    ),
    "analyze": (
        "Pure offline analysis (no network calls) of the extracted metadata against "
        "known Fabric **Direct Lake** constraints. Recommends Direct Lake or Import "
        "mode per cube, with itemized reasons for any fallback (e.g. parent-child "
        "hierarchies, semi-additive measures, many-to-many relationships)."
    ),
    "generate": (
        "Pure offline generation (no network calls) of a TMDL semantic model folder "
        "(tables, columns, relationships, DAX measures translated from MDX, "
        "hierarchies) plus starter PySpark notebook scripts. MDX calculated members/"
        "KPIs/parent-child are NOT auto-translated - they're flagged for manual DAX "
        "authoring instead of risking a silently wrong translation."
    ),
    "report": (
        "Generates `MIGRATION_REPORT.md`: a plain-language summary of exactly what "
        "was converted automatically vs. what needs manual attention, with suggested "
        "alternatives for everything not converted. Review this before moving to "
        "Phase 2."
    ),
    "deploy-lake": (
        "Creates the target Fabric Lakehouse (or reuses it if the name already "
        "exists) and patches the generated TMDL's `expressions.tmdl` with its real "
        "SQL analytics endpoint, so the semantic model knows where to find its data."
    ),
    "migrate-data": (
        "Extracts the star-schema tables directly from the on-prem SQL Server and "
        "writes them as Delta tables into the Lakehouse via OneLake, in one step. "
        "Requires this machine to reach BOTH the SQL Server AND Fabric/OneLake at "
        "the same time."
    ),
    "upload-data": (
        "Reads Delta tables already produced locally (Step 1 below) and pushes them "
        "into the Lakehouse via OneLake. Requires ONLY Fabric connectivity from this "
        "machine - no SQL Server access needed here."
    ),
    "deploy-model": (
        "Creates (or updates) the Semantic Model item in the target Fabric workspace "
        "from the generated TMDL folder. If a previously-deployed model's storage "
        "mode needs to flip (Direct Lake <-> Import), Fabric doesn't support that "
        "in-place - this step automatically deletes and recreates the item instead. "
        "Afterwards it also triggers a refresh and waits for it to finish, so the "
        "model is immediately ready for reports - without this, Fabric shows stale "
        "or missing data until someone manually clicks 'Refresh now' in the portal. "
        "For an Import-mode model's very first deploy, this refresh is expected to "
        "fail with a warning (not an error) until you bind connection credentials "
        "once in the Fabric portal - the model itself still deploys successfully; "
        f"see README {_readme_link('step-7', 'Step 7')}."
    ),
}

ALT_DATA_OPTIONS = [
    ("1. Fabric Mirroring for SQL Server (CDC)", "Native, near real-time, zero-ETL - but requires an On-premises Data Gateway (live connectivity) and SQL Server CDC enabled."),
    ("2. Fabric Mirroring for SQL Server 2025 (Change Feed)", "Same idea as #1, log-based instead of CDC tables - requires SQL Server 2025 + Azure Arc enrollment + gateway."),
    ("3. Fabric Open Mirroring", "You push Parquet/CSV + a metadata file to a Fabric landing zone; Fabric auto-converts to Delta with upserts. No gateway required, but you build the extraction/incremental logic yourself."),
    ("4. Fabric Data Factory (Copy Data activity)", "Low-code scheduled pipeline via an On-premises Data Gateway; supports incremental copy via a watermark column."),
    ("5. Fabric Dataflows Gen2", "Power Query-based low-code ingestion; simplest for business-user-maintained transforms on smaller tables; still needs a gateway."),
    ("6. Fabric Notebook (PySpark/JDBC)", "Full custom Spark logic for large tables/complex transforms. This tool already generates a starter script for this path (Step 3's notebooks/*.py)."),
]

PHASE1_STEPS = ["extract", "analyze", "generate", "report"]


def _report_as_docx(markdown_text: str, title: str) -> bytes:
    """Wraps report_export.markdown_to_docx_bytes with a friendly error if
    python-docx isn't installed in this UI environment yet."""
    from ssas_fabric_migrator.ui.report_export import markdown_to_docx_bytes

    return markdown_to_docx_bytes(markdown_text, title)


def _feasibility_as_excel(feasibility_report: dict) -> bytes:
    """Wraps report_export.feasibility_json_to_excel_bytes with a friendly
    error if openpyxl isn't installed in this UI environment yet."""
    from ssas_fabric_migrator.ui.report_export import feasibility_json_to_excel_bytes

    return feasibility_json_to_excel_bytes(feasibility_report)


def _default_python_exe() -> str:
    """Best-effort default for the pipeline "Python executable" field.

    Prefers a sibling `.venv\\Scripts\\python.exe` (the pipeline-dependency
    environment set up per the README's Quickstart) over `sys.executable`,
    since the latter is *this UI's own* interpreter (typically `.venv-ui`)
    and almost never has pyodbc/pyarrow/pythonnet/deltalake installed -
    using it by mistake is a common source of `ModuleNotFoundError` on the
    very first pipeline step.
    """
    candidate = os.path.join(REPO_ROOT, ".venv", "Scripts", "python.exe")
    if os.path.isfile(candidate):
        return candidate
    return sys.executable


def init_state():
    st.session_state.setdefault("env_file", os.path.join("config", ".env"))
    st.session_state.setdefault("output_dir", "output")
    st.session_state.setdefault("lakehouse_mode", "existing")
    st.session_state.setdefault("lakehouse_name", "RetailLakehouse")
    st.session_state.setdefault("available_lakehouses", [])
    st.session_state.setdefault("local_export_dir", "output\\delta")
    st.session_state.setdefault("local_delta_dir", "output\\delta")
    st.session_state.setdefault("migration_method", "direct")
    st.session_state.setdefault("python_exe", _default_python_exe())
    st.session_state.setdefault("env_values", {k: d for k, _, _, d in ENV_FIELDS})

    # Auto-load real values from an existing env file on the very first run of
    # each session, so features that read st.session_state["env_values"]
    # directly (list_lakehouses, run_local_export) work immediately without
    # requiring a manual "Load from file" click first - unlike the
    # orchestrator-backed pipeline steps, which already re-read the env file
    # fresh from disk on every run regardless of UI state.
    if not st.session_state.get("_env_autoloaded"):
        st.session_state["_env_autoloaded"] = True
        full_path = os.path.join(REPO_ROOT, st.session_state["env_file"])
        if os.path.exists(full_path):
            try:
                load_env_file(st.session_state["env_file"])
            except Exception:
                pass  # leave placeholder defaults; "Load from file" remains available manually


def load_env_file(path: str):
    from ssas_fabric_migrator.cli.orchestrator import load_env

    full_path = path if os.path.isabs(path) else os.path.join(REPO_ROOT, path)
    env = load_env(full_path)
    for key, _, _, _ in ENV_FIELDS:
        if key in env:
            st.session_state["env_values"][key] = env[key]


def save_env_file(path: str):
    full_path = path if os.path.isabs(path) else os.path.join(REPO_ROOT, path)
    os.makedirs(os.path.dirname(full_path), exist_ok=True)
    with open(full_path, "w", encoding="utf-8") as f:
        for key, _, _, _ in ENV_FIELDS:
            f.write(f"{key}={st.session_state['env_values'].get(key, '')}\n")


def _render_log(log_area, text):
    """
    Renders subprocess output in a fixed-height, scrollable box instead of
    st.code()'s default (long lines force a horizontal scrollbar, which is
    awkward for wide tracebacks/paths). Wraps long lines instead
    (white-space: pre-wrap + word-break) and scrolls vertically once the
    content is taller than the box - no horizontal scrolling needed.
    """
    escaped = html.escape(text)
    log_area.markdown(
        '<div style="max-height: 420px; overflow-y: auto; overflow-x: hidden; '
        'white-space: pre-wrap; word-break: break-word; font-family: '
        '\'Source Code Pro\', monospace; font-size: 0.85em; line-height: 1.4; '
        'padding: 0.75em 1em; border-radius: 0.5rem; border: 1px solid rgba(128,128,128,0.3); '
        f'background-color: rgba(128,128,128,0.1);">{escaped}</div>',
        unsafe_allow_html=True,
    )


def _stream_subprocess(cmd, log_area):
    log_lines = [f"$ {' '.join(cmd)}", ""]
    _render_log(log_area, "\n".join(log_lines))
    with st.spinner("Running... this may take a while depending on the step."):
        proc = subprocess.Popen(
            cmd, cwd=REPO_ROOT, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            text=True, bufsize=1,
        )
        for line in proc.stdout:  # type: ignore[union-attr]
            log_lines.append(line.rstrip("\n"))
            _render_log(log_area, "\n".join(log_lines))
        proc.wait()
    return proc.returncode


def run_steps(steps: list[str], log_area):
    """Run the given orchestrator steps as a subprocess, streaming output live."""
    cmd = [
        st.session_state["python_exe"], "-m", "ssas_fabric_migrator.cli.orchestrator",
        "--steps", ",".join(steps),
        "--env-file", st.session_state["env_file"],
        "--output-dir", st.session_state["output_dir"],
        "--lakehouse-name", st.session_state["lakehouse_name"],
        "--semantic-model-name", st.session_state["env_values"].get("SEMANTIC_MODEL_NAME", "RetailCubeDemo"),
    ]
    if "upload-data" in steps:
        cmd += ["--local-delta-dir", st.session_state["local_delta_dir"]]

    returncode = _stream_subprocess(cmd, log_area)
    if returncode == 0:
        st.success(f"Steps [{', '.join(steps)}] completed successfully.")
    else:
        st.error(f"Steps [{', '.join(steps)}] failed with exit code {returncode}. See log above.")


def run_local_export(log_area):
    """Runs loader.py --target local directly (this step isn't part of the orchestrator's
    ALL_STEPS - it's an optional Phase 1 export used only by the offline-transfer path)."""
    out = st.session_state["output_dir"]
    metadata_path = os.path.join(out, "cube_metadata.json")
    ev = st.session_state["env_values"]
    cmd = [
        st.session_state["python_exe"], "-m", "ssas_fabric_migrator.datamover.loader",
        "--metadata", metadata_path,
        "--sql-server", ev.get("SQL_SERVER", ""),
        "--sql-database", ev.get("SQL_DATABASE", ""),
        "--target", "local",
        "--output-dir", st.session_state["local_export_dir"],
    ]
    returncode = _stream_subprocess(cmd, log_area)
    if returncode == 0:
        st.success("Local export completed successfully.")
    else:
        st.error(f"Local export failed with exit code {returncode}. See log above.")


def list_lakehouses():
    """Lists existing Lakehouse items in the configured workspace, via a one-off subprocess
    call using the pipeline's Python interpreter (keeps requests/azure-identity out of the
    UI's own isolated venv)."""
    ev = st.session_state["env_values"]
    script = (
        "import json, os\n"
        "from ssas_fabric_migrator.deploy.fabric_client import FabricClient\n"
        "c = FabricClient(os.environ['FABRIC_TENANT_ID'], os.environ['FABRIC_CLIENT_ID'], os.environ['FABRIC_CLIENT_SECRET'])\n"
        "print(json.dumps(c.list_items(os.environ['FABRIC_WORKSPACE_ID'], 'Lakehouse')))\n"
    )
    env = dict(os.environ)
    for key in ("FABRIC_TENANT_ID", "FABRIC_CLIENT_ID", "FABRIC_CLIENT_SECRET", "FABRIC_WORKSPACE_ID"):
        env[key] = ev.get(key, "")
    result = subprocess.run(
        [st.session_state["python_exe"], "-c", script],
        cwd=REPO_ROOT, capture_output=True, text=True, env=env,
    )
    if result.returncode != 0:
        raise RuntimeError(result.stderr.strip() or "Unknown error listing Lakehouses")
    return json.loads(result.stdout.strip().splitlines()[-1])


def _rewrite_readme_anchors(markdown_text: str) -> str:
    """Rewrites in-file anchor links like "(#3-prerequisites)" into absolute links to the
    README on GitHub, where they actually resolve (GitHub auto-generates heading anchors;
    Streamlit's markdown renderer does not, and this tab only shows excerpts of the file)."""
    return re.sub(r"\]\(#([a-z0-9\-]+)\)", rf"]({README_GITHUB_URL}#\1)", markdown_text)


def render_readme_tab():
    st.subheader("What this tool is - quick summary")
    st.markdown(
        "A code-based accelerator that migrates an on-premises **SQL Server Analysis "
        "Services (SSAS) Multidimensional** cube to **Microsoft Fabric**, producing a "
        "Direct Lake (or Import, with documented reasons) Power BI semantic model "
        "backed by Delta tables in a Fabric Lakehouse. It runs in two phases: "
        "**Phase 1 (on-prem)** extracts/analyzes/generates/reports with no Fabric "
        "connectivity, and **Phase 2 (Fabric-connected)** deploys the Lakehouse, "
        "migrates data, and deploys the semantic model with no SSAS connectivity - "
        "so the two halves can be run on entirely separate machines/networks."
    )

    if not os.path.exists(README_PATH):
        st.warning(f"README.md not found at {README_PATH}")
        return

    with open(README_PATH, "r", encoding="utf-8") as f:
        content = f.read()
    sections = {}
    current_title, current_lines = None, []
    for line in content.splitlines():
        if line.startswith("## "):
            if current_title:
                sections[current_title] = "\n".join(current_lines).strip()
            current_title = line[3:].strip()
            current_lines = []
        else:
            current_lines.append(line)
    if current_title:
        sections[current_title] = "\n".join(current_lines).strip()

    purpose_key = next((k for k in sections if k.startswith("1.")), None)
    limitations_key = next((k for k in sections if k.startswith("10.")), None)

    st.divider()
    st.markdown(f"### What this tool DOES ({_readme_link('1', 'Section 1: Purpose / Objective')})")
    if purpose_key:
        st.markdown(_rewrite_readme_anchors(sections[purpose_key]))
    else:
        st.info("Purpose/Objective section not found in README.md.")

    st.divider()
    st.markdown(f"### What this tool CANNOT do / Limitations ({_readme_link('10', 'Section 10')})")
    if limitations_key:
        st.markdown(_rewrite_readme_anchors(sections[limitations_key]))
    else:
        st.info("Limitations section not found in README.md.")

    st.divider()
    st.caption(
        "This is a live summary pulled directly from README.md in the repo root. Links to "
        "other sections (e.g. 'Section 3') open the full README on GitHub, since this tab "
        "only shows excerpts - see the full file there for the step-by-step walkthrough, "
        "prerequisites, architecture diagram, and all other sections."
    )
    st.link_button("Open full README.md on GitHub", README_GITHUB_URL, type="primary")


def render_config_tab():
    st.subheader("Connection configuration")

    st.caption(
        "These values are written to a local `.env` file (git-ignored) and read by "
        "every pipeline step, exactly like the CLI's `--env-file` argument."
    )

    col1, col2 = st.columns([3, 1])
    with col1:
        st.session_state["env_file"] = st.text_input(
            "Env file path (relative to repo root, or absolute)",
            value=st.session_state["env_file"],
        )
    with col2:
        st.write("")
        st.write("")
        if st.button("Load from file", type="primary"):
            try:
                load_env_file(st.session_state["env_file"])
                st.success("Loaded.")
            except Exception as e:
                st.error(f"Could not load: {e}")

    env_path_for_display = st.session_state["env_file"]
    full_env_path = (
        env_path_for_display if os.path.isabs(env_path_for_display)
        else os.path.join(REPO_ROOT, env_path_for_display)
    )
    exists_note = "exists" if os.path.exists(full_env_path) else "does not exist yet - will be created on Save"
    st.info(
        f"This env file lives on disk **relative to wherever you cloned/downloaded this "
        f"repository** on this machine (`{REPO_ROOT}` on this particular machine right "
        f"now) - it will resolve to a different folder automatically on someone else's "
        f"machine or a different clone location:\n\n`{full_env_path}`\n\n"
        f"({exists_note}) - open it directly in a text editor if you'd rather edit it "
        f"outside this UI. It's git-ignored, so its contents (including secrets) are "
        f"never committed to the repo."
    )

    with st.form("env_form"):
        for key, label, is_secret, _default in ENV_FIELDS:
            st.session_state["env_values"][key] = st.text_input(
                label, value=st.session_state["env_values"].get(key, ""),
                type="password" if is_secret else "default", key=f"field_{key}",
            )
        submitted = st.form_submit_button("Save to env file", type="primary")
        if submitted:
            try:
                save_env_file(st.session_state["env_file"])
                st.success(f"Saved to {st.session_state['env_file']}")
            except Exception as e:
                st.error(f"Could not save: {e}")

    st.divider()
    st.subheader("Run settings")
    st.caption(
        "These are fixed for this deployment (not saved to the env file) - shown here for "
        "reference only."
    )
    st.text_input(
        "Output directory (relative to repo root)",
        value=st.session_state["output_dir"],
        disabled=True,
        help=(
            "All pipeline artifacts (cube_metadata.json, feasibility_report.json, "
            "SemanticModel\\, MIGRATION_REPORT.md, notebooks\\) are written here."
        ),
    )
    st.text_input(
        "Python executable used to run pipeline steps (auto-detected)",
        value=st.session_state["python_exe"],
        disabled=True,
        help=(
            "Auto-detected as <repo_root>\\.venv\\Scripts\\python.exe - the x64 interpreter "
            f"with pythonnet/pandas/pyarrow/deltalake installed per the README {_readme_link('quickstart', 'Quickstart')}. "
            "Every pipeline step subprocess runs with this interpreter."
        ),
    )


def _render_markdown_report_view(path: str, title: str, docx_filename: str, key_prefix: str, not_found_message: str):
    """Shared 'view + Download as Word' block for a generated Markdown report -
    used by both the Reports tab and (per-step) the Phase 1 tab."""
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            text = f.read()
        try:
            st.download_button(
                "\U0001F4C4 Download as Word (.docx)",
                data=_report_as_docx(text, title),
                file_name=docx_filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key=f"{key_prefix}_docx",
            )
        except ImportError:
            st.warning(
                "Install `python-docx` in the UI environment to enable Word export: "
                "`.venv-ui\\Scripts\\python.exe -m pip install python-docx`"
            )
        st.markdown(text)
    else:
        st.info(not_found_message)


def _render_feasibility_report_view(path: str, key_prefix: str, not_found_message: str):
    """Shared 'view + Download as Excel' block for feasibility_report.json -
    used by both the Reports tab and (per-step) the Phase 1 tab."""
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            feasibility_report = json.load(f)
        try:
            st.download_button(
                "\U0001F4CA Download as Excel (.xlsx)",
                data=_feasibility_as_excel(feasibility_report),
                file_name="feasibility_report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"{key_prefix}_xlsx",
            )
        except ImportError:
            st.warning(
                "Install `openpyxl` in the UI environment to enable Excel export: "
                "`.venv-ui\\Scripts\\python.exe -m pip install openpyxl`"
            )
        st.json(feasibility_report)
    else:
        st.write(not_found_message)


def render_phase1_tab():
    st.subheader("Phase 1: On-Prem (SSAS + SQL Server, no Fabric connectivity needed)")
    out = st.session_state["output_dir"]
    full_out = out if os.path.isabs(out) else os.path.join(REPO_ROOT, out)
    feasibility_path = os.path.join(full_out, "feasibility_report.json")
    report_path = os.path.join(full_out, "MIGRATION_REPORT.md")

    for step in PHASE1_STEPS:
        st.markdown(f"**{STEP_LABELS[step]}**")
        st.caption(STEP_DESCRIPTIONS[step])
        run_clicked = st.button("Run this step", key=f"p1_{step}", type="primary")
        log_area = st.empty()
        if run_clicked:
            run_steps([step], log_area)
        if step == "analyze":
            with st.expander("View / download feasibility_report.json"):
                _render_feasibility_report_view(
                    feasibility_path, key_prefix="p1_feasibility",
                    not_found_message="Not generated yet - run this step first.",
                )
        if step == "report":
            with st.expander("View / download MIGRATION_REPORT.md"):
                _render_markdown_report_view(
                    report_path, "Migration Conversion Report", "MIGRATION_REPORT.docx",
                    key_prefix="p1_migration_report",
                    not_found_message="Not generated yet - run this step first.",
                )
        st.divider()


def render_phase2_tab():
    st.subheader("Phase 2: Fabric-connected (no SSAS connectivity needed)")

    # --- Step 5: Lakehouse ---
    st.markdown(f"**{STEP_LABELS['deploy-lake']}**")
    st.caption(STEP_DESCRIPTIONS["deploy-lake"])
    st.session_state["lakehouse_mode"] = st.radio(
        "Lakehouse to use",
        options=["existing", "new"],
        format_func=lambda v: "Use an existing Lakehouse" if v == "existing" else "Create a new Lakehouse",
        index=0 if st.session_state["lakehouse_mode"] == "existing" else 1,
        horizontal=True, key="lakehouse_mode_radio",
    )
    if st.session_state["lakehouse_mode"] == "existing":
        col_a, col_b = st.columns([1, 2])
        with col_a:
            if st.button("Refresh list of Lakehouses", type="primary"):
                try:
                    with st.spinner("Listing Lakehouses in the Fabric workspace..."):
                        st.session_state["available_lakehouses"] = list_lakehouses()
                    st.success(f"Found {len(st.session_state['available_lakehouses'])} Lakehouse(s).")
                except Exception as e:
                    st.error(f"Could not list Lakehouses: {e}")
        with col_b:
            names = [lh["displayName"] for lh in st.session_state["available_lakehouses"]]
            if names:
                st.session_state["lakehouse_name"] = st.selectbox("Existing Lakehouse", options=names)
            else:
                st.session_state["lakehouse_name"] = st.text_input(
                    "Existing Lakehouse name (click 'Refresh list' above to pick from a dropdown instead)",
                    value=st.session_state["lakehouse_name"],
                )
    else:
        st.session_state["lakehouse_name"] = st.text_input(
            "New Lakehouse name", value=st.session_state["lakehouse_name"]
        )

    if st.button("Deploy / find Lakehouse", key="p2_deploy_lake", type="primary"):
        run_steps(["deploy-lake"], st.empty())

    st.divider()

    # --- Step 6: Data migration ---
    st.markdown("**6. Migrate data into the Lakehouse**")
    st.session_state["migration_method"] = st.radio(
        "How do you want to migrate data into the Lakehouse?",
        options=["direct", "offline"],
        format_func=lambda v: (
            "Direct migration - this machine reaches both SQL Server and Fabric"
            if v == "direct" else
            "Offline transfer - export locally, then upload separately (for isolated/air-gapped networks)"
        ),
        index=0 if st.session_state["migration_method"] == "direct" else 1,
        key="migration_method_radio",
    )

    if st.session_state["migration_method"] == "direct":
        st.caption(STEP_DESCRIPTIONS["migrate-data"])
        if st.button("Migrate data now", key="p2_migrate_direct", type="primary"):
            run_steps(["migrate-data"], st.empty())
    else:
        st.markdown("**Step 1 of 2: export SQL Server tables to local Delta files**")
        st.caption(
            "Runs on the on-prem side - needs only SQL Server access, no Fabric connectivity. "
            "Produces one local Delta folder per table under the export folder below."
        )
        st.session_state["local_export_dir"] = st.text_input(
            "Local export folder", value=st.session_state["local_export_dir"], key="local_export_dir_input"
        )
        if st.button("Export to local Delta files", key="p2_export_local", type="primary"):
            run_local_export(st.empty())

        st.markdown(
            "*(Manually transfer the export folder above from the on-prem machine to this "
            "Fabric-connected machine, using whatever secure file-transfer method your "
            "organization already approves - USB, internal file share, SCP, etc. Nothing is "
            "uploaded directly into OneLake at this point.)*"
        )

        st.markdown("**Step 2 of 2: upload the local Delta files to the Lakehouse**")
        st.caption(STEP_DESCRIPTIONS["upload-data"])
        st.session_state["local_delta_dir"] = st.text_input(
            "Local Delta folder to upload (from Step 1, after transfer)",
            value=st.session_state["local_delta_dir"], key="local_delta_dir_input",
        )
        if st.button("Upload to Lakehouse", key="p2_upload", type="primary"):
            run_steps(["upload-data"], st.empty())

    with st.expander("Other ways to migrate data into Fabric (not implemented by this tool)"):
        st.caption(
            "This tool's own approach (direct pyodbc + `deltalake` write, above) is simple and "
            "dependency-light, but is a one-shot batch snapshot with no CDC/incremental support. "
            "For production/enterprise scenarios, Fabric offers these more robust, natively-"
            f"supported options instead - see README {_readme_link('7', 'Section 7')} for the "
            "full comparison table:"
        )
        for name, desc in ALT_DATA_OPTIONS:
            st.markdown(f"- **{name}** - {desc}")

    st.divider()

    # --- Step 7: Deploy model ---
    st.markdown(f"**{STEP_LABELS['deploy-model']}**")
    st.caption(STEP_DESCRIPTIONS["deploy-model"])
    if st.button("Deploy semantic model", key="p2_deploy_model", type="primary"):
        run_steps(["deploy-model"], st.empty())


def render_reports_tab():
    st.subheader("Generated artifacts")
    out = st.session_state["output_dir"]
    full_out = out if os.path.isabs(out) else os.path.join(REPO_ROOT, out)

    report_path = os.path.join(full_out, "MIGRATION_REPORT.md")
    feasibility_path = os.path.join(full_out, "feasibility_report.json")
    manual_path = os.path.join(full_out, "MANUAL_TRANSLATION_REQUIRED.md")

    _render_markdown_report_view(
        report_path, "Migration Conversion Report", "MIGRATION_REPORT.docx",
        key_prefix="reports_migration_report",
        not_found_message=f"No report found yet at {report_path}. Run Phase 1 first.",
    )

    with st.expander("feasibility_report.json"):
        _render_feasibility_report_view(
            feasibility_path, key_prefix="reports_feasibility",
            not_found_message="Not generated yet.",
        )

    with st.expander("MANUAL_TRANSLATION_REQUIRED.md"):
        _render_markdown_report_view(
            manual_path, "Manual Translation Required", "MANUAL_TRANSLATION_REQUIRED.docx",
            key_prefix="reports_manual",
            not_found_message="Not generated yet (only produced if the cube has calculated members/KPIs).",
        )


def main():
    st.set_page_config(page_title="SSAS -> Fabric Migration Tool", layout="wide")
    init_state()

    st.title("SSAS Multidimensional -> Microsoft Fabric Migration Tool")
    st.caption(
        "Wraps the existing CLI/orchestrator - see the Read Me tab or README.md for full "
        "step-by-step docs, prerequisites, and limitations."
    )

    with st.sidebar:
        st.header("About this tool")
        st.markdown(
            "- **On-prem steps** (Extract, direct data migration) require this app to run on "
            "a Windows machine with network access to the SSAS instance/SQL Server, "
            "using an account with the SSAS **Server Administrator** role.\n"
            "- **Fabric steps** use the service principal configured in the "
            "Configuration tab (client ID/secret) - it must be a **member/contributor** "
            "of the target workspace.\n"
            "- Import-mode semantic models need a one-time manual credential binding "
            f"in the Fabric portal after `deploy-model` - see README {_readme_link('10', 'Section 10')}."
        )

    tab_readme, tab_config, tab_phase1, tab_phase2, tab_reports = st.tabs(
        ["Read Me", "Configuration", "Phase 1: On-Prem", "Phase 2: Fabric", "Reports"]
    )
    with tab_readme:
        render_readme_tab()
    with tab_config:
        render_config_tab()
    with tab_phase1:
        render_phase1_tab()
    with tab_phase2:
        render_phase2_tab()
    with tab_reports:
        render_reports_tab()


if __name__ == "__main__":
    main()
