# -*- coding: utf-8 -*-
"""
ASCEND — Financial Instrument Matching Tool
EU Horizon Europe — Grant Agreement 101096571
Streamlit web app version
"""

import streamlit as st
from io import BytesIO
from datetime import date

from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ═══════════════════════════════════════════════════════════════════════════════
# KNOWLEDGE BASE  (identical to original)
# ═══════════════════════════════════════════════════════════════════════════════
KB = [
    {"n":"IRIS Smart Cities",        "t":["ped","smart"], "s":"l",  "p":["cons","util","pub"], "a":["NL","FR","NORD"],    "f":["demo"],  "str":["h2020","erdf","muni","priv"]},
    {"n":"ATELIER",                  "t":["ped","smart"], "s":"l",  "p":["cons","ppp","pub"],  "a":["NL","ES","EU"],      "f":["demo"],  "str":["h2020","ppp","muni","priv","erdf"]},
    {"n":"MAKING-CITY",              "t":["ped","smart"], "s":"l",  "p":["cons","pub"],        "a":["NL","NORD","EU"],    "f":["demo"],  "str":["h2020","green","muni","erdf"]},
    {"n":"POCITYF",                  "t":["ped","smart"], "s":"l",  "p":["util","pub","ppp"],  "a":["PT","NL","EU"],      "f":["demo"],  "str":["h2020","epc","muni","erdf"]},
    {"n":"SPARCS",                   "t":["ped","smart"], "s":"l",  "p":["cons","pub"],        "a":["DE","NORD","EU"],    "f":["demo"],  "str":["h2020","investeu","muni","erdf"]},
    {"n":"STARDUST",                 "t":["ped","smart"], "s":"l",  "p":["ppp","pub"],         "a":["ES","IT","NORD"],    "f":["demo"],  "str":["h2020","epc","muni"]},
    {"n":"MatchUp",                  "t":["ped","smart"], "s":"l",  "p":["cons","pub"],        "a":["ES","DE","EU"],      "f":["demo"],  "str":["h2020","muni","erdf"]},
    {"n":"SmartEnCity",              "t":["ped","smart"], "s":"l",  "p":["cons","pub"],        "a":["ES","EST","EU"],     "f":["demo"],  "str":["h2020","muni","erdf"]},
    {"n":"TRIANGULUM",               "t":["ped","smart"], "s":"l",  "p":["ppp","priv"],        "a":["EU","NORD"],         "f":["demo"],  "str":["h2020","priv","muni"]},
    {"n":"CityxChange",              "t":["ped","smart"], "s":"l",  "p":["cons","pub"],        "a":["NORD","EU"],         "f":["demo"],  "str":["h2020","muni","priv"]},
    {"n":"REACT",                    "t":["cer","resi"],  "s":"m",  "p":["ppp","pub"],         "a":["ES","IT","EU"],      "f":["demo"],  "str":["h2020","ppp","erdf"]},
    {"n":"COME RES",                 "t":["cer"],         "s":"m",  "p":["pub","coop"],        "a":["ES","IT","EU"],      "f":["rd"],    "str":["h2020","erdf","fit"]},
    {"n":"ComER",                    "t":["cer"],         "s":"xs", "p":["coop"],              "a":["IT"],                "f":["rd"],    "str":["h2020","crowd","fit"]},
    {"n":"REScoopVPP",               "t":["cer","grid"],  "s":"m",  "p":["coop"],              "a":["EU"],                "f":["rd"],    "str":["h2020","erdf","fit"]},
    {"n":"NEON",                     "t":["cer"],         "s":"m",  "p":["cons"],              "a":["EU"],                "f":["rd"],    "str":["h2020","greenb"]},
    {"n":"eCREW",                    "t":["grid"],        "s":"m",  "p":["cons","sme"],        "a":["EU"],                "f":["rd"],    "str":["h2020","greenb","crowd","ppp"]},
    {"n":"Flexigrid",                "t":["grid"],        "s":"m",  "p":["cons","util"],       "a":["ES","IT","EU"],      "f":["demo"],  "str":["h2020","erdf"]},
    {"n":"4RinEU",                   "t":["resi","comm"], "s":"m",  "p":["cons","ppp"],        "a":["EU"],                "f":["demo"],  "str":["h2020","ppp","elena","epc"]},
    {"n":"SPHERE",                   "t":["resi","comm"], "s":"m",  "p":["cons","uni"],        "a":["EU"],                "f":["rd"],    "str":["h2020","green","epc"]},
    {"n":"SUPER-HEERO",              "t":["comm"],        "s":"xs", "p":["sme","cons"],        "a":["EU"],                "f":["rd"],    "str":["h2020","crowd","epc"]},
    {"n":"eNeuron",                  "t":["peb","resi"],  "s":"m",  "p":["cons"],              "a":["EU"],                "f":["rd"],    "str":["greenb","eib","tax"]},
    {"n":"Smart-BEEjS",              "t":["ped"],         "s":"m",  "p":["cons","uni"],        "a":["EU"],                "f":["idea"],  "str":["h2020","green","elena","investeu"]},
    {"n":"REGEN",                    "t":["regen"],       "s":"m",  "p":["uni","pub"],         "a":["AT","BE","EU"],      "f":["rd"],    "str":["heur","muni"]},
    {"n":"ReGreeneration",           "t":["regen"],       "s":"l",  "p":["cons","pub"],        "a":["EU","FR","ES"],      "f":["demo"],  "str":["heur","muni"]},
    {"n":"GINNGER",                  "t":["regen"],       "s":"m",  "p":["cons","pub"],        "a":["ES","EU","IT","FR"], "f":["demo"],  "str":["heur","muni"]},
    {"n":"BE-REEL! (Belgium)",       "t":["resi","social"],"s":"l", "p":["pub","iacp"],        "a":["BE","EU"],           "f":["build"], "str":["life","erdf","muni"]},
    {"n":"EEMI (EE Mortgages)",      "t":["resi","peb"],  "s":"xl", "p":["util","priv"],       "a":["EU"],                "f":["ops"],   "str":["life","investeu","greenb","eib"]},
    {"n":"Parco Romana (Milan)",     "t":["ped","regen"], "s":"xl", "p":["priv","ppp","util"], "a":["IT"],                "f":["build"], "str":["priv","ppp","epc","ppa","greenb"]},
    {"n":"Smart City Kalasatama",    "t":["ped","smart"], "s":"xl", "p":["pub","util"],        "a":["NORD"],              "f":["ops"],   "str":["muni","priv","greenb"]},
    {"n":"Quartier Vauban",          "t":["regen","social"],"s":"m","p":["coop","pub"],        "a":["DE","AT"],           "f":["ops"],   "str":["crowd","muni","kfw"]},
    {"n":"Santander Smart Lighting", "t":["smart","comm"],"s":"m",  "p":["pub"],               "a":["ES"],                "f":["build"], "str":["epc","eeef"]},
    {"n":"Gothenburg Green Bonds",   "t":["ped","regen","smart","resi","mob"],"s":"xl","p":["pub"],"a":["NORD"],          "f":["ops"],   "str":["greenb","muni"]},
    {"n":"Region of Madrid S-Bond",  "t":["social","resi","regen"],"s":"xl","p":["pub","iacp"],"a":["ES"],               "f":["ops"],   "str":["socb","muni","erdf"]},
    {"n":"Križevci Sunny Roofs(HR)", "t":["cer","resi"],  "s":"xs", "p":["coop","pub"],        "a":["EST","EU"],          "f":["build"], "str":["crowd","fit","heur"]},
]

INSTR = {
    "heur":    {"n":"Horizon Europe Grant",                    "tag":"EU",   "vol":"€3M–€20M · 100% eligible costs"},
    "h2020":   {"n":"Horizon 2020/Europe Grant (CL5)",         "tag":"EU",   "vol":"€5M–€25M · 100% eligible costs"},
    "life":    {"n":"LIFE – Clean Energy Transition",          "tag":"EU",   "vol":"€1M–€8M · up to 75% of costs"},
    "erdf":    {"n":"ERDF / EU Structural Funds",              "tag":"EU",   "vol":"€2M–€80M · 50%–85% of costs"},
    "rrf":     {"n":"RRF / National Recovery Plan",            "tag":"EU",   "vol":"Variable by country"},
    "greenb":  {"n":"Green Bond (ICMA GBP / EU GBS)",          "tag":"MKT",  "vol":"€20M–€500M · below-market rate"},
    "socb":    {"n":"Social/Sustainability Bond (ICMA SBP)",   "tag":"MKT",  "vol":"€20M–€500M · measurable social impact"},
    "green":   {"n":"Green Loan",                              "tag":"MKT",  "vol":"€1M–€50M · preferential bank loan"},
    "investeu":{"n":"InvestEU / EIB Guarantees",               "tag":"EU",   "vol":"3x–5x leverage on private capital"},
    "elena":   {"n":"ELENA Facility (EIB)",                    "tag":"EU",   "vol":"€200K–€3M · technical assistance"},
    "eeef":    {"n":"European Energy Efficiency Fund (eeef)",  "tag":"EU",   "vol":"€2M–€50M · EPC public assets"},
    "eib":     {"n":"EIB / CEB Loans",                         "tag":"PUB",  "vol":"€5M–€200M · concessional rate"},
    "ppp":     {"n":"PPP / Project Finance",                   "tag":"MIX",  "vol":"€5M–€500M · 15–30 year contract"},
    "epc":     {"n":"EPC / ESCo (performance contract)",       "tag":"MKT",  "vol":"€500K–€30M · repaid from savings"},
    "crowd":   {"n":"Crowdfunding / Civic Crowdlending",       "tag":"MKT",  "vol":"€50K–€5M · citizens' capital"},
    "kfw":     {"n":"KfW / CDP / National Development Bank",   "tag":"PUB",  "vol":"€500K–€50M · subsidised loans"},
    "fit":     {"n":"Feed-in Tariff / Incentives (GSE, EEG)",  "tag":"PUB",  "vol":"€/kWh tariff · contractual basis"},
    "tax":     {"n":"National Tax Incentives",                 "tag":"PUB",  "vol":"IT: Ecobonus/Conto Termico · DE: BEG"},
    "muni":    {"n":"Municipal Co-investment",                 "tag":"PUB",  "vol":"5%–30% of total budget"},
    "priv":    {"n":"Equity / Private Capital",                "tag":"PRIV", "vol":"Expected IRR 8%–15%"},
    "ppa":     {"n":"PPA (Power Purchase Agreement)",          "tag":"MKT",  "vol":"10–25 year contract · guaranteed revenue"},
}

WHY = {
    "heur":    "Horizon Europe (Cluster 5 — Climate, Energy & Mobility) is the primary EU instrument for innovative urban energy projects. Open calls in 2025-2026 cover PEDs, smart cities, deep renovation and energy communities. It funds 100% of eligible costs and requires a consortium of at least 3 organisations from 3 EU countries.",
    "h2020":   "Horizon 2020 SCC-01 and Built4People calls funded over 30 European lighthouse projects with grants from €5M to €25M. These calls are now active under Horizon Europe CL5: multi-country consortium, real-scale demonstration, replication in follower cities.",
    "life":    "The LIFE-CET programme funds projects that remove market barriers to the energy transition: one-stop shops for renovation, capacity building, green mortgages, energy communities. Budget 2021-2027: €1 billion. Annual call by CINEA; co-financing ≥25% required.",
    "erdf":    "ERDF finances the replication of demonstrated solutions in EU regions, with emphasis on sustainable urban development (at least 8% of resources reserved for urban areas). EU co-financing rate from 50% to 85%.",
    "rrf":     "The Recovery and Resilience Facility channels funds through National Recovery Plans. Italy has allocated €15 billion for energy efficiency in buildings. Germany, France and Spain include public and social building renovation in their NRRPs.",
    "greenb":  "The Green Bond (ICMA GBP or EU GBS) is the primary market debt instrument for large sustainable infrastructure. Requires: (1) Sustainable Finance Framework; (2) Second Party Opinion (SPO) from Sustainalytics or CICERO; (3) Annual Impact Report with measurable KPIs. Real case: Gothenburg (SE) — world's first city Green Bond in 2013. Cost-effective from €20M upwards.",
    "socb":    "The Social/Sustainability Bond (ICMA SBP/SBG) extends the Green Bond structure to measurable social impact: social housing, healthcare, education, economic inclusion. Structure: ICMA Framework → SPO → Annual Impact Report with social KPIs. Real case: Region of Madrid issues Sustainability Bonds with Sustainalytics SPO rated 'credible and impactful'.",
    "green":   "A Green Loan is a bank loan with preferential conditions tied to measurable sustainability KPIs. CDP, EIB via intermediary banks all have dedicated green transition lines. More flexible than a Green Bond — no rating required.",
    "investeu":"InvestEU mobilises private capital through EIB guarantees. Every euro of public guarantee generates 3-5 euros of private investment. Accessible via CDP (IT), KfW (DE), Bpifrance (FR). Ideal for bankable projects with DSCR ≥ 1.25x.",
    "elena":   "ELENA (European Local Energy Assistance) covers pre-investment costs: energy audits, feasibility studies, design, tendering. Every €1 of ELENA must generate at least €20 of investment. It is the gateway to bankability.",
    "eeef":    "The European Energy Efficiency Fund (eeef) finances EPC projects on public assets: buildings, street lighting, public transport. Provides direct loans and forfaiting facilities. Real case: Santander (ES) — 22,300 LED street lamps, 80% energy savings, €9M eeef forfaiting facility.",
    "eib":     "EIB and CEB loans offer concessional rates for projects with strong climate and social impact. Minimum ticket €25M for direct EIB access; smaller amounts via national development banks (CDP, KfW, ICO, Bpifrance).",
    "ppp":     "The PPP structures the project under a long-term contract (15-30 years) transferring operational risks to the private sector. Requires positive Value for Money analysis and European competitive tender. Optimal for infrastructure with stable cash flows.",
    "epc":     "The EPC finances renovation through energy savings guaranteed by the ESCo. Two models: Guaranteed Savings (public entity borrows, ESCo guarantees the savings to service debt — preferred in public sector) and Shared Savings (ESCo finances and splits savings with client). Real case: Santander smart lighting — 80% savings guaranteed.",
    "crowd":   "Crowdfunding activates citizens' savings and creates local co-ownership. The Crowdgranting model — public entity matches citizens' funds once a target is reached — levers public resources to unlock broader private investment. Real case: Križevci (HR) 'Sunny Roofs' — civic crowdlending for a solar REC.",
    "kfw":     "KfW (DE), CDP (IT), EIB via national banks: subsidised loans for energy renovation and housing. In Germany: BEG covers up to 75% of eligible costs. In Italy: CDP National Energy Efficiency Fund (FNEE).",
    "fit":     "Feed-in Tariffs (GSE in Italy via DM FER, EEG in Germany) guarantee fixed revenue for energy fed into the grid, making the project bankable for senior debt. In Italy: REC incentives via D.Lgs. 199/2021.",
    "tax":     "Italy: Superbonus (65% residual 2025), Ecobonus, Conto Termico 2.0 GSE. Germany: BEG up to 70% of eligible costs. France: MaPrimeRénov'. Tax incentives reduce net investment cost and improve project IRR.",
    "muni":    "Direct municipal co-investment (grant, equity or public land contribution) is an essential commitment signal to attract EU funds and private investors. Typically 5-30% of total budget.",
    "priv":    "Private equity (developer, ESG infrastructure funds, utility) is the backbone of market-based projects. Expected IRR 8-15% for mature PCDE projects. EU grants or public incentives reduce risk and attract a broader range of investors.",
    "ppa":     "A Power Purchase Agreement is a long-term energy offtake contract (10-25 years) guaranteeing stable revenue and making the project bankable without public subsidies. The offtake counterparty must have investment-grade credit.",
}

COND = {
    "heur":    "Consortium ≥3 entities from ≥3 EU countries · Open calls at eu-funding-tenders.europa.eu · TRL 4-8",
    "h2020":   "Check active Horizon Europe CL5-2025/2026 calls · Multi-country consortium mandatory",
    "life":    "Entity registered in EU · Annual LIFE-CET call (CINEA) · Co-financing ≥25% own resources",
    "erdf":    "Project included in regional/national operational programme · Additionality required",
    "rrf":     "Access via national/regional NRRP calls · Pre-defined milestones and targets",
    "greenb":  "Certified green framework (ICMA External Review or EU GBS) · Effective minimum ~€20M · Annual impact reporting",
    "socb":    "ICMA SBP/SBG certified framework · SPO from Sustainalytics/CICERO · Annual social KPI report · Min ~€20M",
    "green":   "Binding environmental KPIs in the contract · Energy performance certificate pre and post",
    "investeu":"Bankable project DSCR ≥1.25x · Access via CDP (IT), KfW (DE), Bpifrance (FR) or EIB",
    "elena":   "Investment portfolio ≥€50M · Beneficiary: public body · Application via EIB",
    "eeef":    "EPC on public assets (buildings, lighting, transport) · TA available for tender preparation · Contact via EIB",
    "eib":     "Minimum €25M for direct EIB access · Measurable climate impact · Sovereign or corporate IG guarantee",
    "ppp":     "Positive Value for Money analysis · Risk matrix contractually shared · European competitive tender mandatory",
    "epc":     "Certified energy audit · Qualified ESCo · Performance bond · Verifiable consumption baseline",
    "crowd":   "Platform authorised by national regulator (CONSOB/ESMA) · Max €5M per offering (EU Reg. 2020/1503)",
    "kfw":     "IT: CDP/FNEE standards · DE: KfW Effizienzhaus requirements · Certified energy assessment mandatory",
    "fit":     "IT: GSE registration and DM FER caps · DE: EEG declaration and grid connection",
    "tax":     "IT: CILAS filed, qualified contractor · DE: BEG application before works start · Verify current rates",
    "muni":    "Municipal Council resolution · Three-year public works plan · National borrowing limits apply",
    "priv":    "Business plan with IRR ≥8% · Independent technical due diligence (LTA) · Defined exit strategy",
    "ppa":     "Investment-grade offtake counterparty · Contract ≥10 years for bankability · Defined penalty clauses",
}

TAG_COLORS = {
    "EU":   RGBColor(0x1D, 0x6F, 0xA4),
    "MKT":  RGBColor(0x2E, 0x7D, 0x32),
    "PUB":  RGBColor(0x5D, 0x40, 0x37),
    "PRIV": RGBColor(0x7B, 0x1F, 0xA2),
    "MIX":  RGBColor(0x00, 0x69, 0x64),
}

# ═══════════════════════════════════════════════════════════════════════════════
# MATCHING LOGIC  (identical to original)
# ═══════════════════════════════════════════════════════════════════════════════
def match(tipo, imp_m, paese, prom, fase):
    res = {}

    t_ok = tipo in ["ped","regen","resi","cer","smart","grid","mob","dh","peb","social"]
    p_ok = prom in ["uni","cons","pub","ppp","sme","util"]
    f_ok = fase in ["idea","rd","demo"]
    i_ok = 3 <= imp_m <= 30
    if t_ok and p_ok and f_ok and i_ok: res["heur"] = 1
    elif t_ok and (p_ok or f_ok):       res["heur"] = 2
    elif t_ok:                          res["heur"] = 3

    t_ok = tipo in ["ped","smart","regen","cer","grid","mob","dh"]
    p_ok = prom in ["cons","uni","pub","ppp","util","sme"]
    f_ok = fase in ["rd","demo"]
    if t_ok and p_ok and f_ok:          res["h2020"] = 1
    elif t_ok and (p_ok or f_ok):       res["h2020"] = 2
    elif t_ok:                          res["h2020"] = 3

    t_ok = tipo in ["resi","social","cer","peb","regen","dh","comm"]
    p_ok = prom in ["pub","iacp","coop","cons","uni"]
    i_ok = 1 <= imp_m <= 10
    if t_ok and p_ok and i_ok:          res["life"] = 1
    elif t_ok and (p_ok or i_ok):       res["life"] = 2
    elif t_ok:                          res["life"] = 3

    t_ok = tipo in ["ped","smart","regen","resi","social","cer","mob","dh"]
    p_ok = prom in ["pub","iacp","ppp","cons","coop"]
    f_ok = fase in ["demo","build","ops"]
    if t_ok and p_ok and f_ok:          res["erdf"] = 1
    elif t_ok and p_ok:                 res["erdf"] = 2
    elif t_ok:                          res["erdf"] = 3

    t_ok = tipo in ["resi","social","ped","regen","dh","smart","comm"]
    p_ok = prom in ["pub","iacp","ppp"]
    f_ok = fase in ["build","ops"]
    if t_ok and p_ok and f_ok:          res["rrf"] = 1
    elif t_ok and p_ok:                 res["rrf"] = 2
    elif t_ok:                          res["rrf"] = 3

    t_ok = tipo in ["ped","smart","cer","peb","dh","regen","resi","social","grid","mob"]
    p_ok = prom in ["util","priv","pub","ppp","iacp"]
    f_ok = fase in ["build","ops","demo"]
    if imp_m >= 50 and t_ok and p_ok and f_ok: res["greenb"] = 1
    elif imp_m >= 20 and t_ok and p_ok:        res["greenb"] = 2
    elif imp_m >= 10 and t_ok:                 res["greenb"] = 3

    t_ok = tipo in ["social","resi","regen","ped","smart","mob"]
    p_ok = prom in ["pub","iacp","ppp"]
    f_ok = fase in ["build","ops"]
    if imp_m >= 50 and t_ok and p_ok and f_ok: res["socb"] = 1
    elif imp_m >= 20 and t_ok and p_ok:        res["socb"] = 2
    elif imp_m >= 10 and t_ok:                 res["socb"] = 3

    t_ok = tipo in ["resi","comm","peb","ped","regen","dh","social"]
    p_ok = prom in ["priv","ppp","util","pub","iacp"]
    f_ok = fase in ["build","ops","demo"]
    i_ok = 2 <= imp_m <= 80
    if t_ok and p_ok and f_ok and i_ok:  res["green"] = 1
    elif t_ok and p_ok and i_ok:         res["green"] = 2
    elif t_ok and i_ok:                  res["green"] = 3

    t_ok = tipo in ["ped","smart","regen","resi","social","dh","peb","mob"]
    p_ok = prom in ["pub","iacp","util","ppp","priv"]
    f_ok = fase in ["build","ops"]
    if t_ok and p_ok and f_ok and imp_m >= 10:  res["investeu"] = 1
    elif t_ok and p_ok and imp_m >= 10:          res["investeu"] = 2
    elif t_ok:                                   res["investeu"] = 3

    t_ok = tipo in ["resi","social","ped","regen","dh","comm","smart","peb"]
    p_ok = prom in ["pub","iacp","ppp","cons","util"]
    f_ok = fase in ["idea","rd","demo"]
    if t_ok and p_ok and f_ok:   res["elena"] = 1
    elif t_ok and f_ok:          res["elena"] = 2
    elif f_ok:                   res["elena"] = 3

    t_ok = tipo in ["comm","smart","resi","social","dh","mob"]
    p_ok = prom in ["pub","iacp","util","ppp"]
    f_ok = fase in ["idea","rd","demo","build"]
    if t_ok and p_ok and f_ok and imp_m >= 2:  res["eeef"] = 1
    elif t_ok and p_ok:                         res["eeef"] = 2
    elif t_ok:                                  res["eeef"] = 3

    t_ok = tipo in ["ped","smart","dh","social","regen","mob","grid"]
    p_ok = prom in ["pub","util","ppp","priv","iacp"]
    f_ok = fase in ["build","ops"]
    if t_ok and p_ok and f_ok and imp_m >= 15:  res["eib"] = 1
    elif t_ok and p_ok and imp_m >= 15:          res["eib"] = 2
    elif imp_m >= 15:                            res["eib"] = 3

    t_ok = tipo in ["ped","smart","regen","dh","mob","social","grid"]
    p_ok = prom in ["ppp","priv","util","pub"]
    f_ok = fase in ["build","ops","demo"]
    if t_ok and p_ok and f_ok and imp_m >= 10:        res["ppp"] = 1
    elif t_ok and (p_ok or f_ok) and imp_m >= 10:     res["ppp"] = 2
    elif t_ok and imp_m >= 10:                        res["ppp"] = 3

    t_ok = tipo in ["resi","comm","social","ped","dh","peb"]
    p_ok = prom in ["priv","util","ppp","pub","iacp","sme"]
    f_ok = fase in ["build","ops","demo","rd"]
    if t_ok and p_ok and f_ok:  res["epc"] = 1
    elif t_ok and p_ok:         res["epc"] = 2
    elif t_ok:                  res["epc"] = 3

    if tipo == "cer":                                   res["crowd"] = 1
    elif tipo in ["regen","social"] and prom in ["coop","pub","sme"]: res["crowd"] = 2
    elif tipo in ["resi","peb","smart","ped"] and prom in ["coop","pub","sme"]: res["crowd"] = 3
    if "crowd" in res and imp_m > 20:  res["crowd"] = max(res["crowd"], 3)

    t_ok = tipo in ["resi","social","regen","peb","dh","ped"]
    p_ok = prom in ["pub","iacp","priv","ppp","coop"]
    f_ok = fase in ["build","ops","demo"]
    a_ok = paese in ["DE","AT","IT","FR","BE","NORD"]
    if t_ok and p_ok and f_ok and a_ok:  res["kfw"] = 1
    elif t_ok and p_ok and a_ok:          res["kfw"] = 2
    elif t_ok and a_ok:                   res["kfw"] = 3

    t_ok = tipo in ["cer","grid","peb","ped","dh"]
    f_ok = fase in ["build","ops","demo"]
    if t_ok and f_ok:   res["fit"] = 1
    elif t_ok:          res["fit"] = 2
    else:               res["fit"] = 3

    t_ok = tipo in ["resi","social","peb","comm","regen"]
    a_ok = paese in ["IT","DE","FR","ES","AT","BE"]
    p_ok = prom in ["priv","iacp","pub","coop","ppp"]
    if t_ok and a_ok and p_ok:  res["tax"] = 1
    elif t_ok and a_ok:          res["tax"] = 2
    elif t_ok:                   res["tax"] = 3

    t_ok = tipo in ["ped","regen","smart","social","dh","mob","resi","cer"]
    p_ok = prom in ["pub","iacp","cons","ppp","coop"]
    if t_ok and p_ok:  res["muni"] = 1
    elif t_ok:         res["muni"] = 2
    else:              res["muni"] = 3

    t_ok = tipo in ["ped","smart","peb","comm","dh","mob","regen"]
    p_ok = prom in ["priv","util","ppp","sme"]
    f_ok = fase in ["build","ops"]
    if t_ok and p_ok and f_ok and imp_m >= 5:   res["priv"] = 1
    elif t_ok and p_ok and imp_m >= 5:           res["priv"] = 2
    elif t_ok and imp_m >= 5:                    res["priv"] = 3

    t_ok = tipo in ["cer","grid","peb","ped","dh"]
    p_ok = prom in ["util","priv","ppp","coop","pub"]
    f_ok = fase in ["build","ops","demo"]
    if t_ok and p_ok and f_ok:  res["ppa"] = 1
    elif t_ok and f_ok:          res["ppa"] = 2
    elif t_ok:                   res["ppa"] = 3

    return res


def similar_projects(tipo, imp_m, paese, prom):
    sk = "xs" if imp_m < 5 else "m" if imp_m < 30 else "l" if imp_m < 200 else "xl"
    results = []
    for p in KB:
        sim = 0
        if tipo in p["t"]:                              sim += 40
        if p["s"] == sk:                                sim += 20
        if prom in p["p"]:                              sim += 20
        if paese in p["a"] or "EU" in p["a"]:           sim += 20
        if sim >= 40:
            results.append({"n": p["n"], "sim": sim, "str": p["str"]})
    return sorted(results, key=lambda x: -x["sim"])[:4]


def projects_using(key):
    names = [p["n"] for p in KB if key in p["str"]]
    return ", ".join(names[:3]) + ("…" if len(names) > 3 else "") if names else "—"


# ═══════════════════════════════════════════════════════════════════════════════
# WORD REPORT GENERATION  (identical to original)
# ═══════════════════════════════════════════════════════════════════════════════
def hex_to_rgb(hex_str):
    h = hex_str.lstrip('#')
    return tuple(int(h[i:i+2], 16) for i in (0, 2, 4))

def set_cell_bg(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color.lstrip('#'))
    tcPr.append(shd)

def set_table_no_border(table):
    tbl = table._tbl
    tblPr = tbl.tblPr if tbl.tblPr is not None else OxmlElement('w:tblPr')
    tblBorders = OxmlElement('w:tblBorders')
    for edge in ('top','left','bottom','right','insideH','insideV'):
        tag = OxmlElement(f'w:{edge}')
        tag.set(qn('w:val'), 'none')
        tblBorders.append(tag)
    tblPr.append(tblBorders)

def add_para_spacing(para, before=0, after=0):
    pPr = para._p.get_or_add_pPr()
    spacing = OxmlElement('w:spacing')
    spacing.set(qn('w:before'), str(before))
    spacing.set(qn('w:after'), str(after))
    pPr.append(spacing)


def generate_report(tipo, imp_m, paese, prom, fase,
                    tipo_label, paese_label, prom_label, fase_label, imp_label):
    doc = Document()
    for section in doc.sections:
        section.top_margin    = Cm(2.0)
        section.bottom_margin = Cm(2.0)
        section.left_margin   = Cm(2.5)
        section.right_margin  = Cm(2.5)

    # COVER
    band_table = doc.add_table(rows=1, cols=1)
    set_table_no_border(band_table)
    band_cell = band_table.cell(0, 0)
    set_cell_bg(band_cell, '#1D6FA4')
    p = band_cell.paragraphs[0]; p.text = ""

    doc.add_paragraph()
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p_title.add_run("Financial Instrument Matching Report")
    run.bold = True; run.font.size = Pt(28)
    run.font.color.rgb = RGBColor(0x1D, 0x6F, 0xA4); run.font.name = 'Calibri'
    add_para_spacing(p_title, before=200, after=100)

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p_sub.add_run("ASCEND Project  —  EU Horizon Europe  —  Grant Agreement 101096571")
    run.font.size = Pt(13); run.font.color.rgb = RGBColor(0x4A, 0x55, 0x68); run.font.name = 'Calibri'

    doc.add_paragraph()
    info_table = doc.add_table(rows=3, cols=2)
    set_table_no_border(info_table)
    info_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, (lbl, val) in enumerate([
        ("Date", date.today().strftime("%d %B %Y")),
        ("Country / Region", paese_label),
        ("Investment Amount", imp_label),
    ]):
        lc = info_table.cell(i, 0); rc = info_table.cell(i, 1)
        lr = lc.paragraphs[0].add_run(lbl)
        lr.bold = True; lr.font.size = Pt(11)
        lr.font.color.rgb = RGBColor(0x4A, 0x55, 0x68); lr.font.name = 'Calibri'
        vr = rc.paragraphs[0].add_run(val)
        vr.font.size = Pt(11); vr.font.color.rgb = RGBColor(0x2D, 0x37, 0x48); vr.font.name = 'Calibri'

    doc.add_page_break()

    # SECTION 1
    p_head = doc.add_paragraph()
    run = p_head.add_run("1.  Investment Profile")
    run.bold = True; run.font.size = Pt(15)
    run.font.color.rgb = RGBColor(0x1D, 0x6F, 0xA4); run.font.name = 'Calibri'
    add_para_spacing(p_head, before=120, after=80)

    profile_rows = [
        ("Objective", tipo_label), ("Country / Region", paese_label),
        ("Investment Amount", imp_label), ("Promoter", prom_label), ("Project Phase", fase_label),
    ]
    prof_table = doc.add_table(rows=len(profile_rows)+1, cols=2)
    prof_table.style = 'Table Grid'
    for ci, txt in enumerate(["Field", "Value"]):
        c = prof_table.rows[0].cells[ci]
        set_cell_bg(c, '#1D6FA4')
        run = c.paragraphs[0].add_run(txt)
        run.bold = True; run.font.color.rgb = RGBColor(0xFF,0xFF,0xFF)
        run.font.size = Pt(11); run.font.name = 'Calibri'
    for ri, (lbl, val) in enumerate(profile_rows):
        row = prof_table.rows[ri+1]
        set_cell_bg(row.cells[0], '#E8F2FA')
        lr = row.cells[0].paragraphs[0].add_run(lbl)
        lr.bold = True; lr.font.size = Pt(10); lr.font.name = 'Calibri'
        lr.font.color.rgb = RGBColor(0x2D, 0x37, 0x48)
        vr = row.cells[1].paragraphs[0].add_run(val)
        vr.font.size = Pt(10); vr.font.name = 'Calibri'
        vr.font.color.rgb = RGBColor(0x2D, 0x37, 0x48)

    doc.add_paragraph()

    # SECTION 2
    results = match(tipo, imp_m, paese, prom, fase)
    p_head2 = doc.add_paragraph()
    run = p_head2.add_run("2.  Financial Instruments — Results")
    run.bold = True; run.font.size = Pt(15)
    run.font.color.rgb = RGBColor(0x1D, 0x6F, 0xA4); run.font.name = 'Calibri'
    add_para_spacing(p_head2, before=120, after=80)

    for lv, lv_label, hdr_hex, card_hex in [
        (1, "✅  Strongly Recommended", '#1A6B48', '#E8F8F2'),
        (2, "🔶  Worth Evaluating",     '#7A6200', '#FFF9E6'),
        (3, "ℹ️  Complementary",         '#4A5568', '#F4F5F7'),
    ]:
        keys_in_level = [k for k, v in results.items() if v == lv]
        if not keys_in_level:
            continue
        sec_table = doc.add_table(rows=1, cols=1)
        set_table_no_border(sec_table)
        sec_cell = sec_table.cell(0, 0)
        set_cell_bg(sec_cell, hdr_hex)
        sr = sec_cell.paragraphs[0].add_run(f"  {lv_label}  ({len(keys_in_level)} instruments)")
        sr.bold = True; sr.font.size = Pt(12)
        sr.font.color.rgb = RGBColor(0xFF,0xFF,0xFF); sr.font.name = 'Calibri'

        for key in keys_in_level:
            instr = INSTR.get(key, {})
            card_table = doc.add_table(rows=5, cols=2)
            set_table_no_border(card_table)

            name_cell = card_table.cell(0, 0)
            name_cell.merge(card_table.cell(0, 1))
            set_cell_bg(name_cell, card_hex)
            np_ = name_cell.paragraphs[0]
            nr = np_.add_run(f"  {instr.get('n', key)}")
            nr.bold = True; nr.font.size = Pt(12); nr.font.name = 'Calibri'
            nr.font.color.rgb = RGBColor(0x2D, 0x37, 0x48)
            tr = np_.add_run(f"  [{instr.get('tag','')}]")
            tr.bold = True; tr.font.size = Pt(9); tr.font.name = 'Calibri'
            tr.font.color.rgb = TAG_COLORS.get(instr.get('tag',''), RGBColor(0x4A,0x55,0x68))

            vol_cell = card_table.cell(1, 0)
            vol_cell.merge(card_table.cell(1, 1))
            set_cell_bg(vol_cell, card_hex)
            vp_ = vol_cell.paragraphs[0]
            br = vp_.add_run("  Typical volume: ")
            br.bold = True; br.font.size = Pt(9); br.font.name = 'Calibri'
            br.font.color.rgb = RGBColor(0x2D,0x37,0x48)
            vr2 = vp_.add_run(instr.get('vol',''))
            vr2.font.size = Pt(9); vr2.font.name = 'Calibri'
            hr2, hg2, hb2 = hex_to_rgb(hdr_hex)
            vr2.font.color.rgb = RGBColor(hr2, hg2, hb2)

            desc_cell = card_table.cell(2, 0)
            desc_cell.merge(card_table.cell(2, 1))
            set_cell_bg(desc_cell, '#FFFFFF')
            dr = desc_cell.paragraphs[0].add_run(WHY.get(key, ""))
            dr.font.size = Pt(9); dr.font.name = 'Calibri'
            dr.font.color.rgb = RGBColor(0x2D, 0x37, 0x48)

            proj_cell = card_table.cell(3, 0)
            proj_cell.merge(card_table.cell(3, 1))
            set_cell_bg(proj_cell, '#FFFFFF')
            pp_ = proj_cell.paragraphs[0]
            pr1 = pp_.add_run("  Real projects: ")
            pr1.bold = True; pr1.font.size = Pt(9); pr1.font.name = 'Calibri'
            pr1.font.color.rgb = RGBColor(0x2D, 0x37, 0x48)
            pr2 = pp_.add_run(projects_using(key))
            pr2.font.size = Pt(9); pr2.font.name = 'Calibri'
            pr2.font.color.rgb = RGBColor(0x4A, 0x55, 0x68)

            cond_cell = card_table.cell(4, 0)
            cond_cell.merge(card_table.cell(4, 1))
            set_cell_bg(cond_cell, '#F8F9FA')
            cp_ = cond_cell.paragraphs[0]
            cr1 = cp_.add_run("  Key conditions: ")
            cr1.bold = True; cr1.font.size = Pt(8); cr1.font.name = 'Calibri'
            cr1.font.color.rgb = RGBColor(0x4A, 0x55, 0x68)
            cr2 = cp_.add_run(COND.get(key, ""))
            cr2.font.size = Pt(8); cr2.font.name = 'Calibri'
            cr2.font.color.rgb = RGBColor(0x4A, 0x55, 0x68)

            doc.add_paragraph()

    # SECTION 3
    doc.add_paragraph()
    p_head3 = doc.add_paragraph()
    run = p_head3.add_run("3.  Recommended Financing Structure")
    run.bold = True; run.font.size = Pt(15)
    run.font.color.rgb = RGBColor(0x1D, 0x6F, 0xA4); run.font.name = 'Calibri'
    add_para_spacing(p_head3, before=120, after=80)

    lv1_keys = [k for k, v in results.items() if v == 1]
    lv2_keys = [k for k, v in results.items() if v == 2]
    struct_text = "Based on the project profile analysed, the recommended financing structure combines:\n\n"
    if lv1_keys:
        struct_text += f"Primary instruments (Tier 1): {', '.join(INSTR[k]['n'] for k in lv1_keys if k in INSTR)}.\n\n"
    if lv2_keys:
        struct_text += f"Complementary instruments (Tier 2): {', '.join(INSTR[k]['n'] for k in lv2_keys if k in INSTR)}.\n\n"
    struct_text += ("A typical blended finance structure combines EU grants (30-40% of total cost) "
                    "with concessional public loans (30-40%) and private co-investment or market debt (20-40%). "
                    "The exact mix depends on the bankability of the project, risk allocation and the availability "
                    "of national incentives in the target country.")

    box_table = doc.add_table(rows=1, cols=2)
    set_table_no_border(box_table)
    stripe_cell = box_table.cell(0, 0)
    set_cell_bg(stripe_cell, '#1D6FA4')
    stripe_cell.width = Cm(0.4)
    stripe_cell.paragraphs[0].text = " "
    text_cell = box_table.cell(0, 1)
    set_cell_bg(text_cell, '#E8F2FA')
    tr_ = text_cell.paragraphs[0].add_run(struct_text)
    tr_.font.size = Pt(10); tr_.font.name = 'Calibri'
    tr_.font.color.rgb = RGBColor(0x2D, 0x37, 0x48)

    doc.add_paragraph()

    # SECTION 4
    p_head4 = doc.add_paragraph()
    run = p_head4.add_run("4.  Similar Real Projects from the ASCEND Knowledge Base")
    run.bold = True; run.font.size = Pt(15)
    run.font.color.rgb = RGBColor(0x1D, 0x6F, 0xA4); run.font.name = 'Calibri'
    add_para_spacing(p_head4, before=120, after=80)

    sim_projects = similar_projects(tipo, imp_m, paese, prom)
    if sim_projects:
        sim_table = doc.add_table(rows=len(sim_projects)+1, cols=3)
        sim_table.style = 'Table Grid'
        for ci, h in enumerate(["Project Name", "Similarity", "Key Instruments Used"]):
            c = sim_table.rows[0].cells[ci]
            set_cell_bg(c, '#1D6FA4')
            r = c.paragraphs[0].add_run(h)
            r.bold = True; r.font.color.rgb = RGBColor(0xFF,0xFF,0xFF)
            r.font.size = Pt(10); r.font.name = 'Calibri'
        for ri, proj in enumerate(sim_projects):
            row = sim_table.rows[ri+1]
            bg = '#E8F2FA' if ri % 2 == 0 else '#FFFFFF'
            for c in row.cells: set_cell_bg(c, bg)
            r0 = row.cells[0].paragraphs[0].add_run(proj["n"])
            r0.font.name = 'Calibri'; r0.font.size = Pt(10)
            r1 = row.cells[1].paragraphs[0].add_run(f"{proj['sim']}%")
            r1.font.name = 'Calibri'; r1.font.size = Pt(10)
            row.cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            instrs_used = ", ".join(INSTR[s]["n"] if s in INSTR else s for s in proj["str"][:3])
            r2 = row.cells[2].paragraphs[0].add_run(instrs_used)
            r2.font.name = 'Calibri'; r2.font.size = Pt(9)

    doc.add_paragraph()
    doc.add_paragraph()

    # DISCLAIMER
    disc_table = doc.add_table(rows=1, cols=1)
    set_table_no_border(disc_table)
    dc = disc_table.cell(0, 0)
    set_cell_bg(dc, '#F4F5F7')
    dr = dc.paragraphs[0].add_run(
        "DISCLAIMER — This report is generated automatically by the ASCEND Financial Instrument Matching Tool "
        "(EU Horizon Europe — Grant Agreement 101096571). It is intended for informational and orientation purposes only "
        "and does not constitute financial, legal or investment advice. The matching results are based on a simplified "
        "scoring model applied to a reference knowledge base of 34 European projects. Always consult a qualified "
        "financial advisor before making investment or funding decisions. Generated on "
        + date.today().strftime("%d %B %Y") + "."
    )
    dr.font.size = Pt(7); dr.font.name = 'Calibri'
    dr.font.color.rgb = RGBColor(0x4A, 0x55, 0x68)

    return doc


# ═══════════════════════════════════════════════════════════════════════════════
# DROPDOWN OPTIONS
# ═══════════════════════════════════════════════════════════════════════════════
TIPO_OPTS = [
    ("Positive Energy District (PED)",            "ped"),
    ("Urban neighbourhood regeneration",           "regen"),
    ("Deep renovation of residential buildings",   "resi"),
    ("Social / affordable housing",                "social"),
    ("Commercial building renovation",             "comm"),
    ("Renewable Energy Community (REC/CER)",       "cer"),
    ("Plus Energy Building (PEB)",                 "peb"),
    ("District heating / cooling",                 "dh"),
    ("Smart City / integrated urban district",     "smart"),
    ("Smart grid / demand flexibility",            "grid"),
    ("Sustainable urban mobility",                 "mob"),
]

PAESE_OPTS = [
    ("Italy",                              "IT"),
    ("Germany",                            "DE"),
    ("France",                             "FR"),
    ("Spain",                              "ES"),
    ("Netherlands",                        "NL"),
    ("Portugal",                           "PT"),
    ("Austria",                            "AT"),
    ("Belgium",                            "BE"),
    ("Northern Europe (NO,SE,FI,DK)",      "NORD"),
    ("Eastern Europe (PL,CZ,HU,RO,BG,HR)","EST"),
    ("Multi-country European",             "EU"),
]

PROM_OPTS = [
    ("Municipality / local public authority",              "pub"),
    ("Public housing authority (IACP / Housing assoc.)",   "iacp"),
    ("Private developer / investor",                       "priv"),
    ("Utility / energy operator",                          "util"),
    ("Public-private consortium (PPP)",                    "ppp"),
    ("Citizens' cooperative / energy community",           "coop"),
    ("University / research centre",                       "uni"),
    ("SME / innovative startup",                           "sme"),
    ("Multi-entity European consortium",                   "cons"),
]

FASE_OPTS = [
    ("Idea / feasibility study",       "idea"),
    ("Design / R&D",                   "rd"),
    ("Pilot / demonstration",          "demo"),
    ("Construction / implementation",  "build"),
    ("Operational / management",       "ops"),
]

UNIT_OPTS = ["Million €", "Thousand €", "Billion €"]


# ═══════════════════════════════════════════════════════════════════════════════
# STREAMLIT UI
# ═══════════════════════════════════════════════════════════════════════════════
def label_to_value(opts, selected_label):
    for label, value in opts:
        if label == selected_label:
            return value
    return opts[0][1]


def main():
    st.set_page_config(
        page_title="ASCEND Financial Instrument Matching Tool",
        page_icon="⚡",
        layout="centered",
    )

    # Header
    st.markdown("""
    <div style="background-color:#1D6FA4; padding:18px 24px; border-radius:8px; margin-bottom:24px;">
        <span style="color:white; font-size:22px; font-weight:bold;">ASCEND</span>
        <span style="color:#B8D8F0; font-size:13px; margin-left:12px;">
            Financial Instrument Matching Tool &nbsp;·&nbsp; EU Horizon Europe 101096571
        </span>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("### Project Profile")
    st.caption("Fill in the 5 fields below to receive personalised financial instrument recommendations.")

    tipo_labels  = [o[0] for o in TIPO_OPTS]
    paese_labels = [o[0] for o in PAESE_OPTS]
    prom_labels  = [o[0] for o in PROM_OPTS]
    fase_labels  = [o[0] for o in FASE_OPTS]

    tipo_label  = st.selectbox("1.  What do you want to do?",    tipo_labels)
    paese_label = st.selectbox("2.  Where?",                     paese_labels)

    col1, col2 = st.columns([2, 2])
    with col1:
        amt_raw = st.number_input("3.  Investment amount", min_value=0.0, value=10.0, step=1.0)
    with col2:
        unit = st.selectbox("Unit", UNIT_OPTS, label_visibility="hidden")

    prom_label = st.selectbox("4.  Who is the promoter?",       prom_labels)
    fase_label = st.selectbox("5.  Current project phase?",     fase_labels)

    st.markdown("---")

    if st.button("⚡  Generate Report", type="primary", use_container_width=True):
        tipo_val  = label_to_value(TIPO_OPTS, tipo_label)
        paese_val = label_to_value(PAESE_OPTS, paese_label)
        prom_val  = label_to_value(PROM_OPTS, prom_label)
        fase_val  = label_to_value(FASE_OPTS, fase_label)

        if unit == "Thousand €":
            imp_m = amt_raw / 1000
        elif unit == "Billion €":
            imp_m = amt_raw * 1000
        else:
            imp_m = amt_raw

        imp_label = f"{amt_raw:,.1f} {unit}"

        with st.spinner("Generating report…"):
            doc = generate_report(
                tipo_val, imp_m, paese_val, prom_val, fase_val,
                tipo_label, paese_label, prom_label, fase_label, imp_label
            )
            bio = BytesIO()
            doc.save(bio)
            bio.seek(0)

        filename = f"ASCEND_Report_{paese_val}_{date.today().strftime('%Y%m%d')}.docx"

        st.success("Report generated successfully!")
        st.download_button(
            label="⬇️  Download Report (.docx)",
            data=bio,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )

    st.markdown("""
    <div style="margin-top:40px; padding:10px; background:#F4F5F7;
                border-radius:6px; font-size:11px; color:#4A5568;">
        ASCEND Project — EU Horizon Europe — Grant Agreement 101096571 —
        This tool is for informational purposes only and does not constitute financial advice.
    </div>
    """, unsafe_allow_html=True)


if __name__ == "__main__":
    main()
