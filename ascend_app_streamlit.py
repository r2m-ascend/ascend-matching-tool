# -*- coding: utf-8 -*-
"""
ASCEND — Financial Instrument Matching Tool  —  Streamlit app (REV10, Release 2 web)
EU Horizon Europe — Grant Agreement 101096571

Run locally:   streamlit run ascend_app_streamlit.py
"""
from datetime import date

import pandas as pd
import streamlit as st

import engine as E

# ═══════════════════════════════════════════════════════════════════════════════
# PAGE SETUP
# ═══════════════════════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="ASCEND Financial Instrument Matching Tool",
    page_icon="💶",
    layout="wide",
)

BLUE = "#1D6FA4"
TAG_HEX = {"EU": "#1D6FA4", "MKT": "#2E7D32", "PUB": "#5D4037", "PRIV": "#7B1FA2", "MIX": "#006964"}
TAG_LABEL = {"EU": "EU programme", "MKT": "Market", "PUB": "Public", "PRIV": "Private", "MIX": "Blended"}
LEVELS = [
    (1, "Strongly recommended", "✅", "#1A6B48", "#E8F8F2"),
    (2, "Worth evaluating",     "🔶", "#7A6200", "#FFF9E6"),
    (3, "Complementary",        "ℹ️", "#4A5568", "#F4F5F7"),
]

st.markdown(
    f"""
    <style>
      .block-container {{ padding-top: 3.2rem; max-width: 1200px; }}
      .asc-header {{ background:{BLUE}; color:#fff; padding:18px 24px; border-radius:10px; margin-bottom:18px; }}
      .asc-header h1 {{ color:#fff; font-size:1.7rem; margin:0; padding:0; }}
      .asc-header p  {{ color:#CFE4F4; margin:4px 0 0 0; font-size:0.95rem; }}
      .asc-tier {{ color:#fff; padding:8px 14px; border-radius:8px; font-weight:600; margin:18px 0 8px 0; }}
      .asc-card {{ border-radius:8px; padding:12px 16px; margin-bottom:10px; border:1px solid #E2E8F0; }}
      .asc-card h4 {{ margin:0 0 4px 0; font-size:1.05rem; color:#2D3748; }}
      .asc-tag {{ display:inline-block; color:#fff; font-size:0.72rem; font-weight:700;
                  padding:1px 8px; border-radius:10px; margin-left:8px; vertical-align:middle; }}
      .asc-vol {{ font-size:0.85rem; margin-bottom:6px; }}
      .asc-why {{ font-size:0.9rem; color:#2D3748; background:#fff; padding:8px 10px; border-radius:6px; }}
      .asc-meta {{ font-size:0.82rem; color:#4A5568; margin-top:6px; }}
      .asc-box {{ border-left:6px solid {BLUE}; background:#E8F2FA; padding:12px 16px; border-radius:6px; color:#2D3748; }}
      .asc-disc {{ font-size:0.75rem; color:#4A5568; background:#F4F5F7; padding:10px 14px; border-radius:6px; }}
    </style>
    <div class="asc-header">
      <h1>ASCEND · Financial Instrument Matching Tool</h1>
      <p>EU Horizon Europe · Grant Agreement 101096571 · {E.VERSION} · Knowledge base of {E.KB_LABEL}</p>
    </div>
    """,
    unsafe_allow_html=True,
)


# ═══════════════════════════════════════════════════════════════════════════════
# HELPERS
# ═══════════════════════════════════════════════════════════════════════════════
def labels(opts):
    return [o[0] for o in opts]


def code_of(opts, label):
    return dict(opts)[label]


def to_million(amount, unit):
    if unit == "Thousand €":
        return amount / 1000
    if unit == "Billion €":
        return amount * 1000
    return amount


def similarity_column():
    """Progress bar in ASCEND blue (colour argument only exists in recent Streamlit versions)."""
    try:
        return st.column_config.ProgressColumn(
            "Similarity", min_value=0, max_value=100, format="%d%%")
    except TypeError:
        return st.column_config.ProgressColumn(
            "Similarity", min_value=0, max_value=100, format="%d%%")


def instr_name(key):
    return E.INSTR[key]["n"] if key in E.INSTR else key


# ═══════════════════════════════════════════════════════════════════════════════
# FORM
# ═══════════════════════════════════════════════════════════════════════════════
st.subheader("Project profile")
st.caption("Fill in the five fields (plus the optional ASCEND city) and press **Run matching**.")

with st.container(border=True):
    c1, c2 = st.columns(2)
    with c1:
        tipo_l = st.selectbox("1. What do you want to do?", labels(E.TIPO_OPTS))
        paese_l = st.selectbox("2. Where?", labels(E.PAESE_OPTS))
        a1, a2 = st.columns([3, 2])
        with a1:
            amount = st.number_input("3. Investment amount", min_value=0.0, value=10.0, step=1.0, format="%.1f")
        with a2:
            unit = st.selectbox("Unit", E.UNIT_OPTS)
    with c2:
        prom_l = st.selectbox("4. Who is the promoter?", labels(E.PROM_OPTS))
        fase_l = st.selectbox("5. Current project phase?", labels(E.FASE_OPTS))
        citta_l = st.selectbox(
            "6. ASCEND city (optional)", labels(E.CITTA_OPTS),
            help="If you select an ASCEND city, its own Solution Packages are not shown as comparable cases.",
        )
    submitted = st.button("Run matching", type="primary", width="stretch")

# Profile currently shown in the form
current = dict(
    tipo=code_of(E.TIPO_OPTS, tipo_l), tipo_label=tipo_l,
    paese=code_of(E.PAESE_OPTS, paese_l), paese_label=paese_l,
    prom=code_of(E.PROM_OPTS, prom_l), prom_label=prom_l,
    fase=code_of(E.FASE_OPTS, fase_l), fase_label=fase_l,
    citta=code_of(E.CITTA_OPTS, citta_l), citta_label=citta_l,
    imp_m=to_million(amount, unit), imp_label=f"{amount:,.1f} {unit}",
)

if submitted:
    if amount <= 0:
        st.error("Please enter an investment amount greater than zero.")
        st.stop()
    st.session_state["asc_profile"] = current

# Results are shown only for the profile of the last "Run matching".
# If the user changes any field afterwards, old results are hidden until the next run.
last = st.session_state.get("asc_profile")
prof = last if last == current else None

# ═══════════════════════════════════════════════════════════════════════════════
# RESULTS
# ═══════════════════════════════════════════════════════════════════════════════
if prof:
    results = E.match(prof["tipo"], prof["imp_m"], prof["paese"], prof["prom"], prof["fase"])
    sims = E.similar_projects(prof["tipo"], prof["imp_m"], prof["paese"], prof["prom"], prof["citta"])
    by_level = {lv: [k for k, v in results.items() if v == lv] for lv, *_ in LEVELS}

    st.divider()
    head_l, head_r = st.columns([3, 1])
    with head_l:
        st.subheader("Matching results")
        st.caption(
            f"{prof['tipo_label']} · {prof['paese_label']} · {prof['imp_label']} · "
            f"{prof['prom_label']} · {prof['fase_label']}"
            + (f" · ASCEND city: {prof['citta_label']}" if prof["citta"] else "")
        )
    with head_r:
        docx_bytes = E.report_bytes(
            prof["tipo"], prof["imp_m"], prof["paese"], prof["prom"], prof["fase"],
            prof["tipo_label"], prof["paese_label"], prof["prom_label"], prof["fase_label"], prof["imp_label"],
            prof["citta"], prof["citta_label"] if prof["citta"] else "",
        )
        st.download_button(
            "⬇ Download Word report",
            data=docx_bytes,
            file_name=f"ASCEND_Report_{prof['paese']}_{date.today().strftime('%Y%m%d')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            width="stretch",
        )

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Strongly recommended", len(by_level[1]))
    m2.metric("Worth evaluating", len(by_level[2]))
    m3.metric("Complementary", len(by_level[3]))
    m4.metric("Comparable projects", len(sims))

    tab_instr, tab_struct, tab_sim = st.tabs(
        ["Financial instruments", "Recommended structure", "Similar real projects"]
    )

    # ── Instruments by tier ────────────────────────────────────────────────────
    with tab_instr:
        if not results:
            st.info("No instrument matches this profile.")
        for lv, lv_label, icon, hdr_hex, card_hex in LEVELS:
            keys = by_level[lv]
            if not keys:
                continue
            st.markdown(
                f'<div class="asc-tier" style="background:{hdr_hex}">{icon}&nbsp; {lv_label} '
                f'({len(keys)} instrument{"s" if len(keys) != 1 else ""})</div>',
                unsafe_allow_html=True,
            )
            cols = st.columns(2)
            for i, key in enumerate(keys):
                ins = E.INSTR.get(key, {})
                tag = ins.get("tag", "")
                with cols[i % 2]:
                    st.markdown(
                        f"""
                        <div class="asc-card" style="background:{card_hex}">
                          <h4>{ins.get('n', key)}
                            <span class="asc-tag" style="background:{TAG_HEX.get(tag, '#4A5568')}"
                                  title="{TAG_LABEL.get(tag, '')}">{tag}</span></h4>
                          <div class="asc-vol" style="color:{hdr_hex}"><b style="color:#2D3748">Typical volume:</b> {ins.get('vol', '')}</div>
                          <div class="asc-why">{E.WHY.get(key, '')}</div>
                          <div class="asc-meta"><b>Real projects:</b> {E.projects_using(key, prof['citta'])}</div>
                          <div class="asc-meta"><b>Key conditions:</b> {E.COND.get(key, '')}</div>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )

    # ── Recommended structure ──────────────────────────────────────────────────
    with tab_struct:
        parts = ["Based on the project profile analysed, the recommended financing structure combines:"]
        if by_level[1]:
            parts.append("<b>Primary instruments (Tier 1):</b> " + ", ".join(instr_name(k) for k in by_level[1]) + ".")
        if by_level[2]:
            parts.append("<b>Complementary instruments (Tier 2):</b> " + ", ".join(instr_name(k) for k in by_level[2]) + ".")
        parts.append(
            "A typical blended finance structure combines EU grants (30-40% of total cost) "
            "with concessional public loans (30-40%) and private co-investment or market debt (20-40%). "
            "The exact mix depends on the bankability of the project, risk allocation and the availability "
            "of national incentives in the target country."
        )
        st.markdown('<div class="asc-box">' + "<br><br>".join(parts) + "</div>", unsafe_allow_html=True)

    # ── Similar projects ───────────────────────────────────────────────────────
    with tab_sim:
        if sims:
            df = pd.DataFrame(
                [
                    {
                        "Project": p["n"],
                        "Similarity": p["sim"],
                        "Key instruments used": ", ".join(instr_name(s) for s in p["str"][:3]),
                    }
                    for p in sims
                ]
            )
            st.dataframe(
                df,
                hide_index=True,
                width="stretch",
                column_config={"Similarity": similarity_column()},
            )
            st.caption(
                "Similarity score: same project type 40 · same size class 20 · same promoter type 20 · "
                "same country or EU-wide 20. Only projects of the same type are shown."
            )
        else:
            st.info("No closely matching projects found in the knowledge base for this profile.")
elif last:
    st.info("You changed the project profile. Press **Run matching** to update the results.")
else:
    st.info("Fill in the project profile above and press **Run matching** to see the results.")

# ═══════════════════════════════════════════════════════════════════════════════
# KNOWLEDGE BASE BROWSER
# ═══════════════════════════════════════════════════════════════════════════════
st.divider()
with st.expander(f"Browse the knowledge base ({len(E.KB)} projects)"):
    tipo_names = {c: l for l, c in E.TIPO_OPTS}
    prom_names = {c: l for l, c in E.PROM_OPTS}
    fase_names = {c: l for l, c in E.FASE_OPTS}
    size_names = {"xs": "< €5M", "m": "€5M–30M", "l": "€30M–200M", "xl": "> €200M"}
    kb_df = pd.DataFrame(
        [
            {
                "Project": p["n"],
                "Type": ", ".join(tipo_names.get(t, t) for t in p["t"]),
                "Size": size_names.get(p["s"], p["s"]),
                "Promoter": ", ".join(prom_names.get(x, x) for x in p["p"]),
                "Area": ", ".join(p["a"]),
                "Phase": ", ".join(fase_names.get(f, f) for f in p["f"]),
                "Financing instruments": ", ".join(instr_name(s) for s in p["str"]),
            }
            for p in E.KB
        ]
    )
    flt = st.multiselect("Filter by type", labels(E.TIPO_OPTS))
    if flt:
        codes = {code_of(E.TIPO_OPTS, l) for l in flt}
        mask = [bool(codes & set(p["t"])) for p in E.KB]
        kb_df = kb_df[mask]
    st.dataframe(kb_df, hide_index=True, width="stretch")

st.markdown(
    '<div class="asc-disc">DISCLAIMER — This tool is part of the ASCEND project (EU Horizon Europe — Grant '
    "Agreement 101096571). It is intended for informational and orientation purposes only and does not constitute "
    "financial, legal or investment advice. The matching results are based on a simplified scoring model applied "
    f"to a reference knowledge base of {len(E.KB)} European projects. Always consult a qualified financial advisor "
    "before making investment or funding decisions.</div>",
    unsafe_allow_html=True,
)
