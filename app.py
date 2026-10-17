import base64

import streamlit as st

import mf_core as core


def copy_table_button(dataframe):
    tsv = dataframe.to_csv(sep="\t", index=False, na_rep="")
    payload = base64.b64encode(tsv.encode("utf-8")).decode("ascii")
    st.components.v1.html(
        f"""
        <button id="copy-table" type="button">Copy for Excel</button>
        <span id="copy-status" role="status" aria-live="polite"></span>
        <script>
        const encoded = atob("{payload}");
        const bytes = Uint8Array.from(encoded, character => character.charCodeAt(0));
        const text = new TextDecoder().decode(bytes);
        const button = document.getElementById("copy-table");
        const status = document.getElementById("copy-status");

        button.addEventListener("click", async () => {{
            try {{
                await navigator.clipboard.writeText(text);
                status.textContent = "Copied";
            }} catch {{
                const textarea = document.createElement("textarea");
                textarea.value = text;
                document.body.appendChild(textarea);
                textarea.select();
                const copied = document.execCommand("copy");
                textarea.remove();
                status.textContent = copied ? "Copied" : "Clipboard access blocked";
            }}
        }});
        </script>
        <style>
        body {{ font-family: sans-serif; margin: 0; color: #f3f4f6; }}
        button {{ padding: 0.45rem 0.75rem; border: 1px solid #d1d5db; border-radius: 4px;
              background: #f3f4f6; color: #111827; font: inherit; cursor: pointer; }}
        button:hover {{ background: #ffffff; }}
        span {{ margin-left: 0.5rem; color: #f3f4f6; font-size: 0.9rem; }}
        </style>
        """,
        height=44,
    )


st.set_page_config(page_title="MF Performance Master", page_icon="📊", layout="wide")
st.title("📊 Mutual Fund Performance Master")

if "master" not in st.session_state:
    st.session_state.master = core.load_master()
master = st.session_state.master

tab_build, tab_query = st.tabs(["1. Upload & build master", "2. Query performance"])

# ---------------------------------------------------------------- build
with tab_build:
    st.write("Upload the Excel file of each category (Liquid, Large Cap, ...) downloaded from Mutual Sahi Hai.")
    files = st.file_uploader("Excel files", type=["xlsx", "xls"], accept_multiple_files=True)
    mode = st.radio(
        "What should happen to the existing master?",
        ["Add / update (existing schemes get the newer data)", "Replace (start a fresh master)"],
        horizontal=True,
    )

    if st.button("Build master", type="primary", disabled=not files):
        frames = []
        for f in files:
            try:
                df, warns = core.extract_excel(f, f.name)
            except Exception as e:
                st.error(f"{f.name}: could not read ({e})")
                continue
            if df is None:
                st.error(f"{f.name}: {warns[0]}")
                continue
            st.success(f"{f.name}: {len(df)} schemes")
            for w in warns:
                st.warning(f"{f.name}: {w}")
            frames.append(df)

        if frames:
            base = core.load_master().iloc[0:0] if mode.startswith("Replace") else st.session_state.master
            st.session_state.master = core.merge_into_master(base, frames)
            core.save_master(st.session_state.master)
            master = st.session_state.master

    st.subheader(f"Master table ({len(master)} schemes)")
    if master.empty:
        st.info("The master is empty. Upload files and click **Build master**.")
    else:
        st.dataframe(master, use_container_width=True, hide_index=True)
        st.download_button("⬇️ Download master CSV", master.to_csv(index=False).encode(),
                           "mf_master_dataset.csv", "text/csv")

# ---------------------------------------------------------------- query
with tab_query:
    if master.empty:
        st.info("Build the master table first.")
    else:
        c1, c2 = st.columns([3, 2])
        text = c1.text_input("Search fund names",
                             placeholder="e.g. HDFC Liquid, SBI Liquid  (comma separated, partial names work)")
        cats = c2.multiselect("Category / source file", sorted(master["Source File"].dropna().unique()))

        matches = core.search(master, text, cats)
        requested_names = [name.strip() for name in text.split(",") if name.strip()]
        not_found = [name for name in requested_names if core.search(master, name, cats).empty]
        if not_found:
            st.warning("Not found in master table: " + ", ".join(not_found))

        if not matches.empty:
            picked = st.multiselect(
                "Pick specific schemes (optional - leave empty to use all search matches)",
                matches["Scheme Name"].tolist(),
            )
            result = matches[matches["Scheme Name"].isin(picked)] if picked else matches
        else:
            result = matches

        if not text and not cats and not picked:
            st.caption("Showing all schemes - type a name above to narrow down.")
        if not result.empty:
            st.write(f"**{len(result)}** scheme(s)")
            copy_table_button(result)
            st.dataframe(result, use_container_width=True, hide_index=True)
            st.download_button("⬇️ Download result CSV", result.to_csv(index=False).encode(),
                               "filtered_funds_data.csv", "text/csv")
