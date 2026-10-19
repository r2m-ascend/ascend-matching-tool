# -*- coding: utf-8 -*-
"""
ASCEND — Financial Instrument Matching Tool  —  engine (REV10, Release 2 web)
EU Horizon Europe — Grant Agreement 101096571

Matching engine, knowledge base (54 projects) and Word report generator.
Logic and data are identical to the desktop REV9 (ascend_tool_REV9.py);
only the user interface changes (Streamlit, see app.py).
"""
import io
from datetime import date

from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

VERSION = "REV10 · Release 2 (web)"
KB_LABEL = "54 European projects (ASCEND T5.6 DB, 29/09/2026)"

import copy

# ═══════════════════════════════════════════════════════════════════════════════
# KNOWLEDGE BASE — 54 real European projects (REV9, Sept 2026)
# "x" = ASCEND city where the case is NOT shown as comparable
#       (own Solution Package / same site)
# ═══════════════════════════════════════════════════════════════════════════════
KB = [
    # ── From clean DB (ASCEND_T5.6_Financed_Projects_DB_clean, 29/09/2026) ──
    { "n":"IRIS Smart Cities", "t":["ped","smart","mob"], "s":"l", "p":["cons","pub","util"], "a":["NL","FR","NORD"], "f":["demo"], "str":["h2020","muni","priv","greenb"]},
    { "n":"ATELIER", "t":["ped","smart"], "s":"m", "p":["pub","cons","ppp"], "a":["NL","ES"], "f":["demo"], "str":["h2020","muni","ppp"]},
    { "n":"COME RES", "t":["cer"], "s":"xs", "p":["uni","coop","cons"], "a":["EU"], "f":["demo"], "str":["h2020","crowd"]},
    { "n":"eCREW", "t":["cer","grid"], "s":"xs", "p":["uni","coop","cons"], "a":["EU"], "f":["demo"], "str":["h2020","crowd"]},
    { "n":"eNeuron", "t":["grid","cer"], "s":"m", "p":["cons","util","uni"], "a":["IT","PT","EST","NORD"], "f":["demo"], "str":["h2020","priv"]},
    { "n":"Flexigrid", "t":["grid"], "s":"m", "p":["cons","util","uni"], "a":["NORD","ES","EST"], "f":["demo"], "str":["h2020","priv"]},
    { "n":"HIKARI (Lyon Confluence)", "t":["peb","ped"], "s":"l", "p":["ppp","priv"], "a":["FR"], "f":["ops"], "str":["ppp","priv"], "x":"lyon"},
    { "n":"MAKING-CITY", "t":["ped","smart"], "s":"m", "p":["cons","pub"], "a":["NL","NORD"], "f":["demo"], "str":["h2020","muni","epc","ppp","greenb"]},
    { "n":"MAtchUP", "t":["ped","smart","mob"], "s":"m", "p":["pub","cons"], "a":["ES","DE"], "f":["demo"], "str":["h2020","muni","priv"]},
    { "n":"Plus Energy Buildings", "t":["peb","resi"], "s":"m", "p":["priv","iacp"], "a":["AT","NORD","FR"], "f":["build"], "str":["epc","priv","green"]},
    { "n":"POCITYF", "t":["ped","smart"], "s":"m", "p":["util","pub","cons"], "a":["NL","PT"], "f":["demo"], "str":["h2020","priv","greenb"]},
    { "n":"REACT", "t":["cer","grid"], "s":"m", "p":["uni","cons","pub"], "a":["FR","IT","EU"], "f":["demo"], "str":["h2020","crowd","erdf"]},
    { "n":"REScoopVPP", "t":["cer","grid"], "s":"xs", "p":["coop"], "a":["EU"], "f":["demo"], "str":["h2020","crowd"]},
    { "n":"Smart-BEEjS", "t":["ped"], "s":"xs", "p":["uni","cons"], "a":["AT","EU"], "f":["rd"], "str":["h2020","muni"]},
    { "n":"SmartEnCity", "t":["ped","smart","resi"], "s":"m", "p":["cons","pub"], "a":["ES","EST","NORD"], "f":["demo"], "str":["h2020","muni","erdf","crowd"]},
    { "n":"SPARCS", "t":["ped","smart"], "s":"m", "p":["cons","uni","pub"], "a":["NORD","DE"], "f":["demo"], "str":["h2020","muni","priv","epc"]},
    { "n":"STARDUST", "t":["ped","smart","mob"], "s":"m", "p":["cons","pub","ppp"], "a":["ES","NORD","IT"], "f":["demo"], "str":["h2020","muni","greenb","ppp"]},
    { "n":"TIPS4PED", "t":["ped"], "s":"xs", "p":["cons","pub"], "a":["IT","EST","EU"], "f":["demo"], "str":["heur","muni"]},
    { "n":"TRIANGULUM", "t":["ped","smart"], "s":"m", "p":["cons","uni","pub"], "a":["NL","NORD","EU"], "f":["demo"], "str":["h2020","ppp","priv"]},
    { "n":"SPHERE", "t":["resi","comm"], "s":"m", "p":["priv","cons"], "a":["EU"], "f":["demo"], "str":["h2020","priv"]},
    { "n":"4RinEU", "t":["resi"], "s":"xs", "p":["uni","iacp","cons"], "a":["NORD","NL","ES"], "f":["demo"], "str":["h2020","epc"]},
    { "n":"NEON", "t":["cer"], "s":"m", "p":["cons","coop"], "a":["FR","ES","IT","PT"], "f":["demo"], "str":["heur","crowd","epc"]},
    { "n":"REGEN", "t":["regen"], "s":"m", "p":["cons","pub"], "a":["IT","ES","EU"], "f":["demo"], "str":["heur","muni","priv"]},
    { "n":"SUPER-HEERO", "t":["comm","evc"], "s":"xs", "p":["sme","priv"], "a":["IT"], "f":["ops"], "str":["crowd","h2020","epc"]},
    { "n":"Lyon Mobility Hub (SP4)", "t":["log","mob"], "s":"xs", "p":["ppp","pub"], "a":["FR"], "f":["build"], "str":["ppp","heur","priv"], "x":"lyon"},
    { "n":"Munich Harthof Mobility Stations (SP4)", "t":["mob","social"], "s":"xs", "p":["iacp","pub"], "a":["DE"], "f":["build"], "str":["muni"], "x":"munich"},
    { "n":"Decade Energy (depot charging + BESS)", "t":["evc","log"], "s":"m", "p":["priv","sme"], "a":["FR","EU"], "f":["build"], "str":["eaas","priv"]},
    { "n":"Rouen low-emission bus fleet", "t":["mob"], "s":"l", "p":["pub"], "a":["FR"], "f":["build"], "str":["eib","muni"]},
    { "n":"Bicing (Barcelona bike sharing)", "t":["mob"], "s":"l", "p":["ppp","pub"], "a":["ES"], "f":["ops"], "str":["ppp","muni"]},
    { "n":"Villerouge-Termenès e-car sharing", "t":["mob","social"], "s":"xs", "p":["pub"], "a":["FR"], "f":["ops"], "str":["muni"]},
    { "n":"RESPONSE (Dijon & Turku)", "t":["ped","smart","mob"], "s":"l", "p":["pub","cons"], "a":["FR","NORD"], "f":["demo"], "str":["h2020","eib","greenb","priv"]},
    { "n":"Electra ultra-fast charging", "t":["evc"], "s":"xl", "p":["priv"], "a":["FR","BE","EU"], "f":["build"], "str":["priv","investeu"]},
    { "n":"NEVERFLAT (bidirectional charging)", "t":["evc","grid"], "s":"m", "p":["uni","cons","pub"], "a":["EST","DE","ES"], "f":["demo"], "str":["heur"]},
    { "n":"Alba Iulia public transport – Lot 1", "t":["mob"], "s":"m", "p":["pub"], "a":["EST"], "f":["ops"], "str":["erdf","muni"]},
    { "n":"Alba Iulia public transport – Lot 2", "t":["mob"], "s":"m", "p":["pub"], "a":["EST"], "f":["ops"], "str":["erdf","muni"]},
    { "n":"Alba Iulia e-bus fleet & charging", "t":["mob","evc"], "s":"l", "p":["pub"], "a":["EST"], "f":["ops"], "str":["erdf","rrf","muni"]},
    { "n":"Alba Iulia intelligent transport system", "t":["mob","smart"], "s":"xs", "p":["pub"], "a":["EST"], "f":["ops"], "str":["rrf","muni"]},
    { "n":"EV charging from lampposts (Brașov)", "t":["evc","smart"], "s":"xs", "p":["sme","pub","ppp"], "a":["EST"], "f":["demo"], "str":["eit","ppp"]},
    { "n":"NG SEED (fleet-as-a-service)", "t":["log","evc"], "s":"xs", "p":["sme","priv"], "a":["EST"], "f":["demo"], "str":["eit","eaas"]},
    { "n":"Autonom EV fleet (Romania)", "t":["mob","evc"], "s":"l", "p":["priv"], "a":["EST"], "f":["build"], "str":["eib","eaas","priv"]},
    { "n":"EXPAND-E Cohesion charging network", "t":["evc"], "s":"xl", "p":["priv","util","cons"], "a":["EST","EU"], "f":["ops"], "str":["cef","priv"]},
    # ── Retained from REV8 (not in the clean DB) ──
    { "n":"CityxChange", "t":["ped","smart"], "s":"l", "p":["cons","pub"], "a":["NORD","EU"], "f":["demo"], "str":["h2020","muni","priv"]},
    { "n":"ComER", "t":["cer"], "s":"xs", "p":["coop"], "a":["IT"], "f":["rd"], "str":["h2020","crowd","fit"]},
    { "n":"ReGreeneration", "t":["regen"], "s":"l", "p":["cons","pub"], "a":["EU","FR","ES"], "f":["demo"], "str":["heur","muni"]},
    { "n":"GINNGER", "t":["regen"], "s":"m", "p":["cons","pub"], "a":["ES","EU","IT","FR"], "f":["demo"], "str":["heur","muni"]},
    { "n":"BE-REEL! (Belgium)", "t":["resi","social"], "s":"l", "p":["pub","iacp"], "a":["BE","EU"], "f":["build"], "str":["life","erdf","muni"]},
    { "n":"EEMI (EE Mortgages)", "t":["resi","peb"], "s":"xl", "p":["util","priv"], "a":["EU"], "f":["ops"], "str":["life","investeu","greenb","eib"]},
    { "n":"Parco Romana (Milan)", "t":["ped","regen"], "s":"xl", "p":["priv","ppp","util"], "a":["IT"], "f":["build"], "str":["priv","ppp","epc","ppa","greenb"]},
    { "n":"Smart City Kalasatama", "t":["ped","smart"], "s":"xl", "p":["pub","util"], "a":["NORD"], "f":["ops"], "str":["muni","priv","greenb"]},
    { "n":"Quartier Vauban (FR/DE)", "t":["regen","social"], "s":"m", "p":["coop","pub"], "a":["DE","AT"], "f":["ops"], "str":["crowd","muni","kfw"]},
    { "n":"Santander Smart Lighting", "t":["smart","comm"], "s":"m", "p":["pub"], "a":["ES"], "f":["build"], "str":["epc","eeef"]},
    { "n":"Gothenburg Green Bonds", "t":["ped","regen","smart","resi","mob"], "s":"xl", "p":["pub"], "a":["NORD"], "f":["ops"], "str":["greenb","muni"]},
    { "n":"Region of Madrid S-Bond", "t":["social","resi","regen"], "s":"xl", "p":["pub","iacp"], "a":["ES"], "f":["ops"], "str":["socb","muni","erdf"]},
    { "n":"Križevci Sunny Roofs(HR)", "t":["cer","resi"], "s":"xs", "p":["coop","pub"], "a":["EST","EU"], "f":["build"], "str":["crowd","fit","heur"]},
]

# ═══════════════════════════════════════════════════════════════════════════════
# INSTRUMENT CATALOGUE
# ═══════════════════════════════════════════════════════════════════════════════
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
    "cef":     {"n":"CEF Transport – AFIF (grant + bank loan)", "tag":"EU",   "vol":"EU grant blended with a loan from an implementing partner bank"},
    "eit":     {"n":"EIT Urban Mobility (pilot co-funding)",    "tag":"EU",   "vol":"Small tickets · KB cases €0.4M–€0.7M"},
    "eaas":    {"n":"As-a-Service / Operational leasing",       "tag":"MKT",  "vol":"Zero upfront CAPEX · long-term service fee"},
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
    "cef":     "The Connecting Europe Facility (Transport) funds publicly accessible recharging and refuelling points through the Alternative Fuels Infrastructure Facility (AFIF). It is a blending instrument: the EU grant is only paid if the project also has financing from an implementing partner (EIB, a national promotional bank or a commercial bank). Eligible sites are on the TEN-T network and in urban nodes. Real case: EXPAND-E Cohesion, over 1,000 high-power charging points in nine cohesion countries, with the CEF grant blended with private capital and bank debt.",
    "eit":     "EIT Urban Mobility (a Knowledge and Innovation Community supported by Horizon Europe) co-funds city pilots and the market validation of new mobility services through yearly calls. Tickets are small and the typical consortium mixes a city, a company and a research partner. It is a good first step to test a solution before scaling it with ERDF, loans or private capital. Real cases: EV charging from lampposts in Brașov (€0.43M budget) and NG SEED fleet-as-a-service in six Central European cities (€0.66M budget).",
    "eaas":    "In the as-a-service model a specialised operator finances, owns and runs the asset (chargers, batteries, vehicles) and the user pays a fee or subscription, with no upfront CAPEX. The operator raises infrastructure equity and debt against long-term service contracts. Real cases: Decade Energy (depot charging and battery storage for logistics, €22M raised from the Eiffel Transition Infrastructure Fund and SET Ventures) and Autonom in Romania (operational leasing and rental of low-emission fleets, €15M EIB loan).",
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
    "cef":     "Site on TEN-T network or urban node · Financing from an implementing partner bank is a precondition · Check open AFIF cut-off dates on CINEA",
    "eit":     "Yearly calls at eiturbanmobility.eu · Consortium with city and industry partner · Co-funding rate and eligibility set by each call",
    "eaas":    "Long-term service contract with a creditworthy user · Residual value and utilisation risk priced in the fee · Operator with access to infrastructure equity",
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
# MATCHING LOGIC
# ═══════════════════════════════════════════════════════════════════════════════
def match(tipo, imp_m, paese, prom, fase):
    res = {}

    # HORIZON EUROPE
    t_ok = tipo in ["ped","regen","resi","cer","smart","grid","mob","dh","peb","social","evc","log"]
    p_ok = prom in ["uni","cons","pub","ppp","sme","util"]
    f_ok = fase in ["idea","rd","demo"]
    i_ok = 3 <= imp_m <= 30
    if t_ok and p_ok and f_ok and i_ok: res["heur"] = 1
    elif t_ok and (p_ok or f_ok):       res["heur"] = 2
    elif t_ok:                          res["heur"] = 3

    # HORIZON 2020 / EUROPE CL5
    t_ok = tipo in ["ped","smart","regen","cer","grid","mob","dh","evc","log"]
    p_ok = prom in ["cons","uni","pub","ppp","util","sme"]
    f_ok = fase in ["rd","demo"]
    if t_ok and p_ok and f_ok:          res["h2020"] = 1
    elif t_ok and (p_ok or f_ok):       res["h2020"] = 2
    elif t_ok:                          res["h2020"] = 3

    # LIFE-CET
    t_ok = tipo in ["resi","social","cer","peb","regen","dh","comm"]
    p_ok = prom in ["pub","iacp","coop","cons","uni"]
    i_ok = 1 <= imp_m <= 10
    if t_ok and p_ok and i_ok:          res["life"] = 1
    elif t_ok and (p_ok or i_ok):       res["life"] = 2
    elif t_ok:                          res["life"] = 3

    # ERDF
    t_ok = tipo in ["ped","smart","regen","resi","social","cer","mob","dh","evc"]
    p_ok = prom in ["pub","iacp","ppp","cons","coop"]
    f_ok = fase in ["demo","build","ops"]
    if t_ok and p_ok and f_ok:          res["erdf"] = 1
    elif t_ok and p_ok:                 res["erdf"] = 2
    elif t_ok:                          res["erdf"] = 3

    # RRF / PNRR
    t_ok = tipo in ["resi","social","ped","regen","dh","smart","comm","mob","evc"]
    p_ok = prom in ["pub","iacp","ppp"]
    f_ok = fase in ["build","ops"]
    if t_ok and p_ok and f_ok:          res["rrf"] = 1
    elif t_ok and p_ok:                 res["rrf"] = 2
    elif t_ok:                          res["rrf"] = 3

    # GREEN BOND
    t_ok = tipo in ["ped","smart","cer","peb","dh","regen","resi","social","grid","mob","evc"]
    p_ok = prom in ["util","priv","pub","ppp","iacp"]
    f_ok = fase in ["build","ops","demo"]
    if imp_m >= 50 and t_ok and p_ok and f_ok: res["greenb"] = 1
    elif imp_m >= 20 and t_ok and p_ok:        res["greenb"] = 2
    elif imp_m >= 10 and t_ok:                 res["greenb"] = 3

    # SOCIAL / SUSTAINABILITY BOND
    t_ok = tipo in ["social","resi","regen","ped","smart","mob"]
    p_ok = prom in ["pub","iacp","ppp"]
    f_ok = fase in ["build","ops"]
    if imp_m >= 50 and t_ok and p_ok and f_ok: res["socb"] = 1
    elif imp_m >= 20 and t_ok and p_ok:        res["socb"] = 2
    elif imp_m >= 10 and t_ok:                 res["socb"] = 3

    # GREEN LOAN
    t_ok = tipo in ["resi","comm","peb","ped","regen","dh","social"]
    p_ok = prom in ["priv","ppp","util","pub","iacp"]
    f_ok = fase in ["build","ops","demo"]
    i_ok = 2 <= imp_m <= 80
    if t_ok and p_ok and f_ok and i_ok:  res["green"] = 1
    elif t_ok and p_ok and i_ok:         res["green"] = 2
    elif t_ok and i_ok:                  res["green"] = 3

    # INVESTEU
    t_ok = tipo in ["ped","smart","regen","resi","social","dh","peb","mob","evc","log"]
    p_ok = prom in ["pub","iacp","util","ppp","priv"]
    f_ok = fase in ["build","ops"]
    if t_ok and p_ok and f_ok and imp_m >= 10:  res["investeu"] = 1
    elif t_ok and p_ok and imp_m >= 10:          res["investeu"] = 2
    elif t_ok:                                   res["investeu"] = 3

    # ELENA
    t_ok = tipo in ["resi","social","ped","regen","dh","comm","smart","peb"]
    p_ok = prom in ["pub","iacp","ppp","cons","util"]
    f_ok = fase in ["idea","rd","demo"]
    if t_ok and p_ok and f_ok:   res["elena"] = 1
    elif t_ok and f_ok:          res["elena"] = 2
    elif f_ok:                   res["elena"] = 3

    # EEEF
    t_ok = tipo in ["comm","smart","resi","social","dh","mob"]
    p_ok = prom in ["pub","iacp","util","ppp"]
    f_ok = fase in ["idea","rd","demo","build"]
    if t_ok and p_ok and f_ok and imp_m >= 2:  res["eeef"] = 1
    elif t_ok and p_ok:                         res["eeef"] = 2
    elif t_ok:                                  res["eeef"] = 3

    # EIB LOANS
    t_ok = tipo in ["ped","smart","dh","social","regen","mob","grid","evc","log"]
    p_ok = prom in ["pub","util","ppp","priv","iacp"]
    f_ok = fase in ["build","ops"]
    if t_ok and p_ok and f_ok and imp_m >= 15:  res["eib"] = 1
    elif t_ok and p_ok and imp_m >= 15:          res["eib"] = 2
    elif imp_m >= 15:                            res["eib"] = 3

    # PPP
    t_ok = tipo in ["ped","smart","regen","dh","mob","social","grid","evc","log"]
    p_ok = prom in ["ppp","priv","util","pub"]
    f_ok = fase in ["build","ops","demo"]
    if t_ok and p_ok and f_ok and imp_m >= 10:        res["ppp"] = 1
    elif t_ok and (p_ok or f_ok) and imp_m >= 10:     res["ppp"] = 2
    elif t_ok and imp_m >= 10:                        res["ppp"] = 3

    # EPC / ESCo
    t_ok = tipo in ["resi","comm","social","ped","dh","peb"]
    p_ok = prom in ["priv","util","ppp","pub","iacp","sme"]
    f_ok = fase in ["build","ops","demo","rd"]
    if t_ok and p_ok and f_ok:  res["epc"] = 1
    elif t_ok and p_ok:         res["epc"] = 2
    elif t_ok:                  res["epc"] = 3

    # CROWDFUNDING
    if tipo == "cer":                                   res["crowd"] = 1
    elif tipo in ["regen","social"] and prom in ["coop","pub","sme"]: res["crowd"] = 2
    elif tipo in ["resi","peb","smart","ped"] and prom in ["coop","pub","sme"]: res["crowd"] = 3
    if "crowd" in res and imp_m > 20:  res["crowd"] = max(res["crowd"], 3)

    # KfW / CDP
    t_ok = tipo in ["resi","social","regen","peb","dh","ped"]
    p_ok = prom in ["pub","iacp","priv","ppp","coop"]
    f_ok = fase in ["build","ops","demo"]
    a_ok = paese in ["DE","AT","IT","FR","BE","NORD"]
    if t_ok and p_ok and f_ok and a_ok:  res["kfw"] = 1
    elif t_ok and p_ok and a_ok:          res["kfw"] = 2
    elif t_ok and a_ok:                   res["kfw"] = 3

    # FIT / FEED-IN
    t_ok = tipo in ["cer","grid","peb","ped","dh"]
    f_ok = fase in ["build","ops","demo"]
    if t_ok and f_ok:   res["fit"] = 1
    elif t_ok:          res["fit"] = 2
    else:               res["fit"] = 3

    # TAX INCENTIVES
    t_ok = tipo in ["resi","social","peb","comm","regen"]
    a_ok = paese in ["IT","DE","FR","ES","AT","BE"]
    p_ok = prom in ["priv","iacp","pub","coop","ppp"]
    if t_ok and a_ok and p_ok:  res["tax"] = 1
    elif t_ok and a_ok:          res["tax"] = 2
    elif t_ok:                   res["tax"] = 3

    # MUNICIPAL CO-INVESTMENT
    t_ok = tipo in ["ped","regen","smart","social","dh","mob","resi","cer","evc","log"]
    p_ok = prom in ["pub","iacp","cons","ppp","coop"]
    if t_ok and p_ok:  res["muni"] = 1
    elif t_ok:         res["muni"] = 2
    else:              res["muni"] = 3

    # PRIVATE EQUITY
    t_ok = tipo in ["ped","smart","peb","comm","dh","mob","regen","evc","log"]
    p_ok = prom in ["priv","util","ppp","sme"]
    f_ok = fase in ["build","ops"]
    if t_ok and p_ok and f_ok and imp_m >= 5:   res["priv"] = 1
    elif t_ok and p_ok and imp_m >= 5:           res["priv"] = 2
    elif t_ok and imp_m >= 5:                    res["priv"] = 3

    # PPA
    t_ok = tipo in ["cer","grid","peb","ped","dh"]
    p_ok = prom in ["util","priv","ppp","coop","pub"]
    f_ok = fase in ["build","ops","demo"]
    if t_ok and p_ok and f_ok:  res["ppa"] = 1
    elif t_ok and f_ok:          res["ppa"] = 2
    elif t_ok:                   res["ppa"] = 3

    # CEF TRANSPORT – AFIF (publicly accessible recharging / refuelling)
    t_ok = tipo == "evc"
    t_weak = tipo in ["mob","log"]
    p_ok = prom in ["priv","util","ppp","pub","cons","sme"]
    f_ok = fase in ["build","demo"]
    if t_ok and p_ok and f_ok:          res["cef"] = 1
    elif t_ok and p_ok:                 res["cef"] = 2
    elif t_ok or (t_weak and f_ok):     res["cef"] = 3

    # EIT URBAN MOBILITY — pilots and early market validation
    t_ok = tipo in ["mob","evc","log","smart"]
    p_ok = prom in ["pub","sme","priv","cons","uni"]
    f_ok = fase in ["idea","rd","demo"]
    i_ok = imp_m <= 2
    if t_ok and p_ok and f_ok and i_ok: res["eit"] = 1
    elif t_ok and f_ok:                 res["eit"] = 2
    elif t_ok and i_ok:                 res["eit"] = 3

    # AS-A-SERVICE / OPERATIONAL LEASING
    t_ok = tipo in ["evc","log","mob"]
    t_weak = tipo in ["comm","peb","dh"]
    p_ok = prom in ["priv","sme","util","pub","iacp","ppp"]
    f_ok = fase in ["build","ops"]
    if t_ok and p_ok and f_ok:          res["eaas"] = 1
    elif t_ok and p_ok:                 res["eaas"] = 2
    elif t_ok or t_weak:                res["eaas"] = 3

    return res


def similar_projects(tipo, imp_m, paese, prom, citta=""):
    sk = "xs" if imp_m < 5 else "m" if imp_m < 30 else "l" if imp_m < 200 else "xl"
    results = []
    for p in KB:
        if citta and p.get("x") == citta:
            continue                      # no circular reference to the city's own case
        if tipo not in p["t"]:
            continue                      # type must match
        sim = 0
        if tipo in p["t"]:                              sim += 40
        if p["s"] == sk:                                sim += 20
        if prom in p["p"]:                              sim += 20
        if paese in p["a"] or "EU" in p["a"]:           sim += 20
        if sim >= 40:
            results.append({"n": p["n"], "sim": sim, "str": p["str"]})
    return sorted(results, key=lambda x: -x["sim"])[:4]


# ═══════════════════════════════════════════════════════════════════════════════
# WORD REPORT GENERATION
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

def set_cell_border(cell, **kwargs):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for edge in ('top','left','bottom','right'):
        if edge in kwargs:
            tag = OxmlElement(f'w:{edge}')
            tag.set(qn('w:val'), kwargs[edge].get('val','single'))
            tag.set(qn('w:sz'), str(kwargs[edge].get('sz', 4)))
            tag.set(qn('w:space'), '0')
            tag.set(qn('w:color'), kwargs[edge].get('color','000000'))
            tcBorders.append(tag)
    tcPr.append(tcBorders)

def add_para_spacing(para, before=0, after=0):
    pPr = para._p.get_or_add_pPr()
    spacing = OxmlElement('w:spacing')
    spacing.set(qn('w:before'), str(before))
    spacing.set(qn('w:after'), str(after))
    pPr.append(spacing)

def add_run_font(run, font_name='Calibri'):
    rPr = run._r.get_or_add_rPr()
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'), font_name)
    rFonts.set(qn('w:hAnsi'), font_name)
    rPr.insert(0, rFonts)

def set_table_no_border(table):
    tbl = table._tbl
    tblPr = tbl.tblPr if tbl.tblPr is not None else OxmlElement('w:tblPr')
    tblBorders = OxmlElement('w:tblBorders')
    for edge in ('top','left','bottom','right','insideH','insideV'):
        tag = OxmlElement(f'w:{edge}')
        tag.set(qn('w:val'), 'none')
        tblBorders.append(tag)
    tblPr.append(tblBorders)

def projects_using(key, citta=""):
    names = [p["n"] for p in KB if key in p["str"] and not (citta and p.get("x") == citta)]
    return ", ".join(names[:3]) + ("…" if len(names) > 3 else "") if names else "—"


def generate_report(tipo, imp_m, paese, prom, fase,
                    tipo_label, paese_label, prom_label, fase_label, imp_label,
                    citta="", citta_label=""):
    doc = Document()

    # ── Page margins ────────────────────────────────────────────────────────────
    for section in doc.sections:
        section.top_margin    = Cm(2.0)
        section.bottom_margin = Cm(2.0)
        section.left_margin   = Cm(2.5)
        section.right_margin  = Cm(2.5)

    # ══════════════════════════════════════════════════════
    # PAGE 1 — COVER
    # ══════════════════════════════════════════════════════

    # Top colour band
    band_table = doc.add_table(rows=1, cols=1)
    set_table_no_border(band_table)
    band_cell = band_table.cell(0, 0)
    set_cell_bg(band_cell, '#1D6FA4')
    band_cell.height = Cm(1.2)
    p = band_cell.paragraphs[0]
    p.text = ""
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(0)

    doc.add_paragraph()

    # Main title
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p_title.add_run("Financial Instrument Matching Report")
    run.bold = True
    run.font.size = Pt(28)
    run.font.color.rgb = RGBColor(0x1D, 0x6F, 0xA4)
    run.font.name = 'Calibri'
    add_para_spacing(p_title, before=200, after=100)

    # Subtitle
    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p_sub.add_run("ASCEND Project  —  EU Horizon Europe  —  Grant Agreement 101096571")
    run.font.size = Pt(13)
    run.font.color.rgb = RGBColor(0x4A, 0x55, 0x68)
    run.font.name = 'Calibri'
    add_para_spacing(p_sub, before=60, after=60)

    # Separator line
    sep_table = doc.add_table(rows=1, cols=1)
    set_table_no_border(sep_table)
    sep_cell = sep_table.cell(0, 0)
    set_cell_bg(sep_cell, '#E8F2FA')
    p = sep_cell.paragraphs[0]
    p.text = ""
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(0)

    doc.add_paragraph()

    # Date and key info box
    info_table = doc.add_table(rows=3, cols=2)
    set_table_no_border(info_table)
    info_table.alignment = WD_TABLE_ALIGNMENT.CENTER

    info_data = [
        ("Date", date.today().strftime("%d %B %Y")),
        ("Country / Region", paese_label),
        ("Investment Amount", imp_label),
    ]
    for i, (lbl, val) in enumerate(info_data):
        lc = info_table.cell(i, 0)
        rc = info_table.cell(i, 1)
        lp = lc.paragraphs[0]
        lr = lp.add_run(lbl)
        lr.bold = True
        lr.font.size = Pt(11)
        lr.font.color.rgb = RGBColor(0x4A, 0x55, 0x68)
        lr.font.name = 'Calibri'
        lc.width = Cm(5)
        vp = rc.paragraphs[0]
        vr = vp.add_run(val)
        vr.font.size = Pt(11)
        vr.font.color.rgb = RGBColor(0x2D, 0x37, 0x48)
        vr.font.name = 'Calibri'

    doc.add_paragraph()
    doc.add_paragraph()

    # Bottom band
    bot_table = doc.add_table(rows=1, cols=1)
    set_table_no_border(bot_table)
    bot_cell = bot_table.cell(0, 0)
    set_cell_bg(bot_cell, '#1D6FA4')
    p = bot_cell.paragraphs[0]
    p.text = ""

    # Page break
    doc.add_page_break()

    # ══════════════════════════════════════════════════════
    # PAGE 2+ — BODY
    # ══════════════════════════════════════════════════════

    # ── 1. Investment Profile ───────────────────────────────────────────────────
    p_head = doc.add_paragraph()
    run = p_head.add_run("1.  Investment Profile")
    run.bold = True
    run.font.size = Pt(15)
    run.font.color.rgb = RGBColor(0x1D, 0x6F, 0xA4)
    run.font.name = 'Calibri'
    add_para_spacing(p_head, before=120, after=80)

    profile_rows = [
        ("Objective",         tipo_label),
        ("Country / Region",  paese_label),
        ("Investment Amount", imp_label),
        ("Promoter",          prom_label),
        ("Project Phase",     fase_label),
    ]
    if citta:
        profile_rows.append(("ASCEND city", citta_label + " (own Solution Packages excluded from comparables)"))
    prof_table = doc.add_table(rows=len(profile_rows)+1, cols=2)
    prof_table.style = 'Table Grid'
    prof_table.alignment = WD_TABLE_ALIGNMENT.LEFT

    hdr_row = prof_table.rows[0]
    for ci, txt in enumerate(["Field", "Value"]):
        c = hdr_row.cells[ci]
        set_cell_bg(c, '#1D6FA4')
        p = c.paragraphs[0]
        run = p.add_run(txt)
        run.bold = True
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        run.font.size = Pt(11)
        run.font.name = 'Calibri'
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT

    for ri, (lbl, val) in enumerate(profile_rows):
        row = prof_table.rows[ri+1]
        lc = row.cells[0]
        vc = row.cells[1]
        set_cell_bg(lc, '#E8F2FA')
        lp = lc.paragraphs[0]
        lr = lp.add_run(lbl)
        lr.bold = True
        lr.font.size = Pt(10)
        lr.font.name = 'Calibri'
        lr.font.color.rgb = RGBColor(0x2D, 0x37, 0x48)
        vp = vc.paragraphs[0]
        vr = vp.add_run(val)
        vr.font.size = Pt(10)
        vr.font.name = 'Calibri'
        vr.font.color.rgb = RGBColor(0x2D, 0x37, 0x48)

    for col_idx, width in enumerate([Cm(5.5), Cm(10)]):
        for cell in prof_table.columns[col_idx].cells:
            cell.width = width

    doc.add_paragraph()

    # ── 2. Results by level ─────────────────────────────────────────────────────
    results = match(tipo, imp_m, paese, prom, fase)

    levels = [
        (1, "✅  Strongly Recommended",  '#1A6B48', '#E8F8F2'),
        (2, "🔶  Worth Evaluating",       '#7A6200', '#FFF9E6'),
        (3, "ℹ️  Complementary",          '#4A5568', '#F4F5F7'),
    ]

    p_head2 = doc.add_paragraph()
    run = p_head2.add_run("2.  Financial Instruments — Results")
    run.bold = True
    run.font.size = Pt(15)
    run.font.color.rgb = RGBColor(0x1D, 0x6F, 0xA4)
    run.font.name = 'Calibri'
    add_para_spacing(p_head2, before=120, after=80)

    for lv, lv_label, hdr_hex, card_hex in levels:
        keys_in_level = [k for k, v in results.items() if v == lv]
        if not keys_in_level:
            continue

        # Section header band
        sec_table = doc.add_table(rows=1, cols=1)
        set_table_no_border(sec_table)
        sec_cell = sec_table.cell(0, 0)
        set_cell_bg(sec_cell, hdr_hex)
        sp = sec_cell.paragraphs[0]
        sr = sp.add_run(f"  {lv_label}  ({len(keys_in_level)} instruments)")
        sr.bold = True
        sr.font.size = Pt(12)
        sr.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        sr.font.name = 'Calibri'
        sp.paragraph_format.space_before = Pt(4)
        sp.paragraph_format.space_after  = Pt(4)

        for key in keys_in_level:
            instr = INSTR.get(key, {})
            instr_name = instr.get("n", key)
            instr_tag  = instr.get("tag", "")
            instr_vol  = instr.get("vol", "")
            why_text   = WHY.get(key, "")
            cond_text  = COND.get(key, "")
            projects   = projects_using(key, citta)

            card_table = doc.add_table(rows=5, cols=2)
            set_table_no_border(card_table)
            card_table.alignment = WD_TABLE_ALIGNMENT.LEFT

            # Row 0: Name + Tag
            name_cell = card_table.cell(0, 0)
            tag_cell  = card_table.cell(0, 1)
            set_cell_bg(name_cell, card_hex)
            set_cell_bg(tag_cell,  card_hex)
            name_cell.merge(tag_cell)
            np_ = name_cell.paragraphs[0]
            nr = np_.add_run(f"  {instr_name}")
            nr.bold = True
            nr.font.size = Pt(12)
            nr.font.name = 'Calibri'
            nr.font.color.rgb = RGBColor(0x2D, 0x37, 0x48)
            tr = np_.add_run(f"  [{instr_tag}]")
            tr.bold = True
            tr.font.size = Pt(9)
            tr.font.name = 'Calibri'
            tag_rgb = TAG_COLORS.get(instr_tag, RGBColor(0x4A, 0x55, 0x68))
            tr.font.color.rgb = tag_rgb
            np_.paragraph_format.space_before = Pt(5)
            np_.paragraph_format.space_after  = Pt(2)

            # Row 1: Typical volume
            vol_cell = card_table.cell(1, 0)
            vol_cell2 = card_table.cell(1, 1)
            set_cell_bg(vol_cell,  card_hex)
            set_cell_bg(vol_cell2, card_hex)
            vol_cell.merge(vol_cell2)
            vp_ = vol_cell.paragraphs[0]
            vp_.add_run("  Typical volume: ").bold = True
            vr2 = vp_.add_run(instr_vol)
            vr2.font.size = Pt(9)
            vr2.font.name = 'Calibri'
            hr, hg, hb = hex_to_rgb(hdr_hex)
            vr2.font.color.rgb = RGBColor(hr, hg, hb)
            vp_.paragraph_format.space_before = Pt(0)
            vp_.paragraph_format.space_after  = Pt(2)
            # Fix bold run
            br = vp_.runs[0]
            br.font.size = Pt(9)
            br.font.name = 'Calibri'
            br.font.color.rgb = RGBColor(0x2D, 0x37, 0x48)

            # Row 2: Description
            desc_cell  = card_table.cell(2, 0)
            desc_cell2 = card_table.cell(2, 1)
            set_cell_bg(desc_cell,  '#FFFFFF')
            set_cell_bg(desc_cell2, '#FFFFFF')
            desc_cell.merge(desc_cell2)
            dp_ = desc_cell.paragraphs[0]
            dp_.add_run("  ").font.size = Pt(1)
            dr = dp_.add_run(why_text)
            dr.font.size = Pt(9)
            dr.font.name = 'Calibri'
            dr.font.color.rgb = RGBColor(0x2D, 0x37, 0x48)
            dp_.paragraph_format.space_before = Pt(3)
            dp_.paragraph_format.space_after  = Pt(3)

            # Row 3: Real projects
            proj_cell  = card_table.cell(3, 0)
            proj_cell2 = card_table.cell(3, 1)
            set_cell_bg(proj_cell,  '#FFFFFF')
            set_cell_bg(proj_cell2, '#FFFFFF')
            proj_cell.merge(proj_cell2)
            pp_ = proj_cell.paragraphs[0]
            pr1 = pp_.add_run("  Real projects: ")
            pr1.bold = True
            pr1.font.size = Pt(9)
            pr1.font.name = 'Calibri'
            pr1.font.color.rgb = RGBColor(0x2D, 0x37, 0x48)
            pr2 = pp_.add_run(projects)
            pr2.font.size = Pt(9)
            pr2.font.name = 'Calibri'
            pr2.font.color.rgb = RGBColor(0x4A, 0x55, 0x68)
            pp_.paragraph_format.space_before = Pt(0)
            pp_.paragraph_format.space_after  = Pt(2)

            # Row 4: Conditions
            cond_cell  = card_table.cell(4, 0)
            cond_cell2 = card_table.cell(4, 1)
            set_cell_bg(cond_cell,  '#F8F9FA')
            set_cell_bg(cond_cell2, '#F8F9FA')
            cond_cell.merge(cond_cell2)
            cp_ = cond_cell.paragraphs[0]
            cr1 = cp_.add_run("  Key conditions: ")
            cr1.bold = True
            cr1.font.size = Pt(8)
            cr1.font.name = 'Calibri'
            cr1.font.color.rgb = RGBColor(0x4A, 0x55, 0x68)
            cr2 = cp_.add_run(cond_text)
            cr2.font.size = Pt(8)
            cr2.font.name = 'Calibri'
            cr2.font.color.rgb = RGBColor(0x4A, 0x55, 0x68)
            cp_.paragraph_format.space_before = Pt(0)
            cp_.paragraph_format.space_after  = Pt(5)

            doc.add_paragraph()

    # ── 3. Recommended financing structure ──────────────────────────────────────
    doc.add_paragraph()
    p_head3 = doc.add_paragraph()
    run = p_head3.add_run("3.  Recommended Financing Structure")
    run.bold = True
    run.font.size = Pt(15)
    run.font.color.rgb = RGBColor(0x1D, 0x6F, 0xA4)
    run.font.name = 'Calibri'
    add_para_spacing(p_head3, before=120, after=80)

    lv1_keys = [k for k, v in results.items() if v == 1]
    lv2_keys = [k for k, v in results.items() if v == 2]

    struct_text = "Based on the project profile analysed, the recommended financing structure combines:\n\n"
    if lv1_keys:
        names_lv1 = ", ".join(INSTR[k]["n"] for k in lv1_keys if k in INSTR)
        struct_text += f"Primary instruments (Tier 1): {names_lv1}.\n\n"
    if lv2_keys:
        names_lv2 = ", ".join(INSTR[k]["n"] for k in lv2_keys if k in INSTR)
        struct_text += f"Complementary instruments (Tier 2): {names_lv2}.\n\n"
    struct_text += (
        "A typical blended finance structure combines EU grants (30-40% of total cost) "
        "with concessional public loans (30-40%) and private co-investment or market debt (20-40%). "
        "The exact mix depends on the bankability of the project, risk allocation and the availability "
        "of national incentives in the target country."
    )

    # Blue-border box
    box_table = doc.add_table(rows=1, cols=2)
    set_table_no_border(box_table)
    stripe_cell = box_table.cell(0, 0)
    set_cell_bg(stripe_cell, '#1D6FA4')
    stripe_cell.width = Cm(0.4)
    p = stripe_cell.paragraphs[0]
    p.text = " "
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(0)

    text_cell = box_table.cell(0, 1)
    set_cell_bg(text_cell, '#E8F2FA')
    tp_ = text_cell.paragraphs[0]
    tr_ = tp_.add_run(struct_text)
    tr_.font.size = Pt(10)
    tr_.font.name = 'Calibri'
    tr_.font.color.rgb = RGBColor(0x2D, 0x37, 0x48)
    tp_.paragraph_format.space_before = Pt(6)
    tp_.paragraph_format.space_after  = Pt(6)

    doc.add_paragraph()

    # ── 4. Similar real projects ─────────────────────────────────────────────────
    p_head4 = doc.add_paragraph()
    run = p_head4.add_run("4.  Similar Real Projects from the ASCEND Knowledge Base")
    run.bold = True
    run.font.size = Pt(15)
    run.font.color.rgb = RGBColor(0x1D, 0x6F, 0xA4)
    run.font.name = 'Calibri'
    add_para_spacing(p_head4, before=120, after=80)

    sim_projects = similar_projects(tipo, imp_m, paese, prom, citta)

    if sim_projects:
        sim_table = doc.add_table(rows=len(sim_projects)+1, cols=3)
        sim_table.style = 'Table Grid'
        headers = ["Project Name", "Similarity", "Key Instruments Used"]
        for ci, h in enumerate(headers):
            c = sim_table.rows[0].cells[ci]
            set_cell_bg(c, '#1D6FA4')
            p = c.paragraphs[0]
            r = p.add_run(h)
            r.bold = True
            r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            r.font.size = Pt(10)
            r.font.name = 'Calibri'

        for ri, proj in enumerate(sim_projects):
            row = sim_table.rows[ri+1]
            bg = '#E8F2FA' if ri % 2 == 0 else '#FFFFFF'
            for c in row.cells:
                set_cell_bg(c, bg)
            row.cells[0].paragraphs[0].add_run(proj["n"]).font.name = 'Calibri'
            row.cells[0].paragraphs[0].runs[0].font.size = Pt(10)
            row.cells[1].paragraphs[0].add_run(f"{proj['sim']}%").font.name = 'Calibri'
            row.cells[1].paragraphs[0].runs[0].font.size = Pt(10)
            row.cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            instrs_used = ", ".join(INSTR[s]["n"] if s in INSTR else s for s in proj["str"][:3])
            row.cells[2].paragraphs[0].add_run(instrs_used).font.name = 'Calibri'
            row.cells[2].paragraphs[0].runs[0].font.size = Pt(9)
    else:
        p_no = doc.add_paragraph("No closely matching projects found in the knowledge base for this profile.")
        p_no.runs[0].font.size = Pt(10)
        p_no.runs[0].font.name = 'Calibri'

    doc.add_paragraph()
    doc.add_paragraph()

    # ── 5. Disclaimer ────────────────────────────────────────────────────────────
    disc_table = doc.add_table(rows=1, cols=1)
    set_table_no_border(disc_table)
    dc = disc_table.cell(0, 0)
    set_cell_bg(dc, '#F4F5F7')
    dp_ = dc.paragraphs[0]
    dr = dp_.add_run(
        "DISCLAIMER — This report is generated automatically by the ASCEND Financial Instrument Matching Tool "
        "(EU Horizon Europe — Grant Agreement 101096571). It is intended for informational and orientation purposes only "
        "and does not constitute financial, legal or investment advice. The matching results are based on a simplified "
        "scoring model applied to a reference knowledge base of 54 European projects (REV10). Always consult a qualified "
        "financial advisor before making investment or funding decisions. Generated on "
        + date.today().strftime("%d %B %Y") + "."
    )
    dr.font.size = Pt(7)
    dr.font.name = 'Calibri'
    dr.font.color.rgb = RGBColor(0x4A, 0x55, 0x68)
    dp_.paragraph_format.space_before = Pt(6)
    dp_.paragraph_format.space_after  = Pt(6)

    return doc


def report_bytes(*args, **kwargs):
    """Build the Word report in memory and return it as bytes (for st.download_button)."""
    doc = generate_report(*args, **kwargs)
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ═══════════════════════════════════════════════════════════════════════════════
# FORM OPTIONS (label, code)
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
    ("Sustainable urban mobility (public transport, sharing)", "mob"),
    ("EV charging infrastructure",                 "evc"),
    ("Urban logistics hub / last mile",            "log"),
]

CITTA_OPTS = [
    ("None / other city", ""),
    ("Lyon",              "lyon"),
    ("Munich",            "munich"),
    ("Budapest",          "budapest"),
    ("Porto",             "porto"),
    ("Prague",            "prague"),
    ("Stockholm",         "stockholm"),
]

PAESE_OPTS = [
    ("Italy",                         "IT"),
    ("Germany",                       "DE"),
    ("France",                        "FR"),
    ("Spain",                         "ES"),
    ("Netherlands",                   "NL"),
    ("Portugal",                      "PT"),
    ("Austria",                       "AT"),
    ("Belgium",                       "BE"),
    ("Northern Europe (NO,SE,FI,DK)", "NORD"),
    ("Eastern Europe (PL,CZ,HU,RO,BG,HR)", "EST"),
    ("Multi-country European",        "EU"),
]

PROM_OPTS = [
    ("Municipality / local public authority",     "pub"),
    ("Public housing authority (IACP / Housing association)", "iacp"),
    ("Private developer / investor",              "priv"),
    ("Utility / energy operator",                 "util"),
    ("Public-private consortium (PPP)",           "ppp"),
    ("Citizens' cooperative / energy community",  "coop"),
    ("University / research centre",              "uni"),
    ("SME / innovative startup",                  "sme"),
    ("Multi-entity European consortium",          "cons"),
]

FASE_OPTS = [
    ("Idea / feasibility study",      "idea"),
    ("Design / R&D",                  "rd"),
    ("Pilot / demonstration",         "demo"),
    ("Construction / implementation", "build"),
    ("Operational / management",      "ops"),
]

UNIT_OPTS = ["Million €", "Thousand €", "Billion €"]
