import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter

FONT = "Arial"
HEADER_FILL = PatternFill("solid", fgColor="1F4E5F")
HEADER_FONT = Font(name=FONT, bold=True, color="FFFFFF", size=10)
TITLE_FONT = Font(name=FONT, bold=True, size=14, color="1F4E5F")
SUBTITLE_FONT = Font(name=FONT, italic=True, size=10, color="555555")
BODY_FONT = Font(name=FONT, size=10)
WRAP = Alignment(wrap_text=True, vertical="top")
WRAP_CENTER = Alignment(wrap_text=True, vertical="center", horizontal="center")
THIN = Side(style="thin", color="CCCCCC")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

wb = openpyxl.Workbook()
wb.remove(wb.active)

def style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = WRAP_CENTER
        cell.border = BORDER

def set_widths(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

def write_table(ws, headers, rows, start_row=1, widths=None):
    for j, h in enumerate(headers, start=1):
        ws.cell(row=start_row, column=j, value=h)
    style_header(ws, start_row, len(headers))
    for i, row_data in enumerate(rows, start=start_row + 1):
        for j, val in enumerate(row_data, start=1):
            cell = ws.cell(row=i, column=j, value=val)
            cell.font = BODY_FONT
            cell.alignment = WRAP
            cell.border = BORDER
    ws.freeze_panes = ws.cell(row=start_row + 1, column=1).coordinate
    if widths:
        set_widths(ws, widths)
    return start_row + len(rows)

# ---------------------------------------------------------------------------
# README
# ---------------------------------------------------------------------------
ws = wb.create_sheet("README")
ws["A1"] = "India Hospital Fire Incident Database (2026– )"
ws["A1"].font = TITLE_FONT
ws["A2"] = "Purpose: track hospital fire incidents in India from 2026 onward to identify recurring causes, analyze implicated products/equipment, and inform hospital-staff training and mitigation measures."
ws["A2"].font = SUBTITLE_FONT
ws["A2"].alignment = WRAP
ws.merge_cells("A2:H2")
ws.row_dimensions[2].height = 30

readme_lines = [
    ("How the workbook is organized", True),
    ("Incidents", "One row per fire incident. This is the primary data-entry sheet."),
    ("Causes_Taxonomy", "Controlled vocabulary for Root_Cause_Category (used as a dropdown in Incidents). Edit/extend as new cause types appear; do not delete codes already referenced by an incident."),
    ("Products_Taxonomy", "Controlled vocabulary for Product_Equipment_Involved (used as a dropdown in Incidents), with common failure modes and the relevant Indian/international standards for each equipment class."),
    ("Failure_Analysis", "Narrative failure-analysis write-ups. Create one row per distinct product/failure-mode cluster once 2+ incidents share a pattern (e.g. 'aged electrical panels in ICUs'), not one per incident."),
    ("Sources", "Every citation backing an Incidents row. One Source_ID per article/report; an incident can cite multiple sources, and one source can cover multiple incidents."),
    ("Training_Recommendations", "Actionable training points derived from causes/failure analyses, tagged by target staff audience."),
    ("Dashboard", "Auto-calculated summary counts/trends pulled from Incidents via formulas. Do not hand-edit; it recalculates as you add rows."),
    ("", ""),
    ("Data entry conventions", True),
    ("Incident_ID", "Format HF-YYYY-NNN (HF = Hospital Fire), sequential within each year, e.g. HF-2026-001."),
    ("Dates", "ISO format YYYY-MM-DD."),
    ("Root_Cause_Category / Product_Equipment_Involved", "Pick from the dropdown. If a new cause or product type appears, add it to the relevant taxonomy sheet FIRST, then select it here."),
    ("Data_Confidence", "Confirmed - Official Report (fire dept/NDMA/court/RTI); Confirmed - Multiple News Sources (2+ independent outlets agree); Single Source - Unverified; Conflicting Reports (sources disagree on key facts - detail the discrepancy in Notes)."),
    ("Source_IDs", "Comma-separated Source_ID values, e.g. S001, S003."),
    ("Severity_Category", "Fatal; Non-Fatal Injury; Near Miss (No Injury); Property/Equipment Damage Only; Unknown. Near misses are IN SCOPE and deliberately logged alongside fatal incidents - a contained fire with no injuries shares the same root cause and is often the same failure mechanism one step from becoming fatal, so it carries equal weight for training and prevention purposes even though it will never make national headlines."),
    ("Scope", "In scope: fires originating in hospital/nursing-home/clinic buildings and their fixed infrastructure (wards, ICUs, oxygen plants, electrical rooms, kitchens) - REGARDLESS of outcome, including near-misses with zero fatalities/injuries. Out of scope for the main log: standalone ambulance/vehicle fires and fires at non-healthcare buildings - log these separately if useful, they are relevant background for product failure analysis (e.g. oxygen cylinders) but are not 'hospital fires'."),
    ("", ""),
    ("Workflow for adding an incident", True),
    ("1.", "Log the source(s) in the Sources sheet first (Source_ID, outlet, URL, dates)."),
    ("2.", "Add a row to Incidents, filling every column you can support with a source; leave uncertain fields blank rather than guessing, and note uncertainty in Notes."),
    ("3.", "If the incident's cause/product matches or extends an existing Failure_Analysis cluster, add its Incident_ID to that row's Linked_Incident_IDs and update the narrative; otherwise wait until a pattern (2+ incidents) emerges."),
    ("4.", "Add or update rows in Training_Recommendations if the incident reveals a new training gap."),
    ("", ""),
    ("Sourcing languages", "Search now deliberately covers Hindi-language regional press (Prabhat Khabar, Amar Ujala, Nav Bharat Live, Haribhoomi, NPG News, Subkuz, Dainik Jagran) in addition to English national/international outlets, since near-miss and smaller incidents are disproportionately covered only by state/regional-language press and are invisible to English-only search. Bengali, Tamil, Telugu, Marathi and Gujarati searches should be run periodically too - a Bengali pass on 2026-09-13 found no hospital-specific incident (only unrelated fires), but that is a negative result for one search session, not proof none occurred; keep retrying as coverage builds. When translating a non-English source, keep the original-language headline/title alongside the English translation in the Sources sheet (see existing S009-S016) so the citation can be independently verified."),
    ("Update log 2026-09-18", "2026-09-18 (window Sep 10-18): 1 new incident (HF-2026-013, MKCG Berhampur, Odisha), 2 updates to existing rows (HF-2026-002 toll reconciled to 6 dead / 17 injured with three arrests; HF-2026-009 follow-up on twin newborns who died after transfer, attributed to prematurity and not counted as fire deaths), new failure analysis FA-003 (AC units) and recommendations TR-008/TR-009. Searched in English, Hindi, Marathi, Odia, Malayalam, Tamil, Gujarati, Bengali, Telugu and Kannada; Bengali, Telugu, Kannada, Tamil and Gujarati returned nothing relevant in the window. Verified negatives: a search hit for a 'Bhind District Hospital fire on Sep 15' was actually an undated Damoh District Hospital article (fan short circuit at the OPD counter, extinguished by staff) and is not logged; a Bhilwara hospital fire and a Baghpat house fire surfaced earlier were out of window or out of scope. Backlog of PRE-WINDOW 2026 fires seen in search results but not yet logged (verify before adding): Dehradun Panacea Hospital ICU (May 20, reportedly one death); Thiruvananthapuram Medical College surgical ICU ventilator (Mar 17); Thoothukudi GH (Mar 24); PMCH Patna microbiology dept (Apr 3); Thuravoor taluk hospital, Alappuzha (Apr 14); Bhilwara Siddhi Vinayak Hospital (Apr 27); Vadodara Anand Hospital (Apr/May); Sri Ganganagar district hospital and a Jodhpur private hospital (~Jun); Lok Bandhu Hospital One Stop Centre, Lucknow (~Jul); Kanpur Hallet Hospital (Aug 7); Karakkulam private hospital, Thiruvananthapuram (Sep 3); and three earlier MKCG Berhampur fires (May 28, Aug 19, Sep 4)."),
    ("Update log 2026-09-28", "2026-09-28 (window Sep 18-28): 2 new incidents - HF-2026-014 (RIMS Adilabad, Telangana, night of Sep 21: window-AC fire in the SNCU, 5 premature infants dead of smoke, inquiry under way) and HF-2026-015 (DMCH Darbhanga, Bihar, Sep 26: MCH-ward electrical panel short circuit, no casualties). FA-002 and FA-003 updated with Adilabad; TR-010 (NICU door egress) and TR-011 (dedicated electrical/AC maintenance ownership) added. Searched English, Hindi, Marathi, Odia, Tamil, Telugu, Kannada, Gujarati, Bengali and Malayalam; only Hindi, Telugu and English produced in-window incidents. Watch items: the Adilabad toll may still rise (further infants were critical) and its six-member committee report was due about Oct 4 - it may name the AC's type, age and maintenance history, which would upgrade FA-003 from hypothesis to evidence. Source cautions: the maintenance claims in S031 are from unnamed sources, ETV Bharat English mis-dates the fire as Sep 22-23, and a Telugu explainer still shows a toll of 3. Pakistani and Bangladeshi hospital fires from the same weeks are out of scope. Additional PRE-WINDOW leads not yet logged (verify before adding): Chhindwara District Hospital, MP (Aug 22, reportedly one death and two injured - seen only in an ETV Bharat Telugu explainer); Jabalpur NSCB Medical College NICU (late Mar, doctor extinguished a fan fire, 27 babies safe); Kolkata Anandalok Hospital, Salt Lake (Apr 21, AC short circuit in an operation theatre, about 100 patients moved); Hangal taluk hospital, Karnataka (Apr, short circuit, medicines burnt); Bidar Guru Nanak Hospital (Apr); Madhavaram primary health centre, Chennai (Jul 2, UPS exploded); Parkar Hospital dialysis unit, Ratnagiri (date not found); Bardhaman government hospital (2026, not checked). The 2026-09-18 backlog list above still stands."),
    ("Initial build", "2026-09-13 - seeded with 12 verified 2026 incidents found via news search (see Sources sheet): 2 fatal (Cuttack ICU; Muzaffarpur ICU) and 1 more fatal found via Hindi-language search (Amravati NICU, HF-2026-012), plus 9 near-misses/minor-injury events with no deaths, 6 of those 9 surfaced only through Hindi-language regional press. This is a starting structure, not a complete dataset - many more near-misses (and likely some fatal events too) go unreported or under-reported in English-language national media, so continued regional-language search plus local/state fire-department and hospital-internal incident logs are the priority next sources to close that gap."),
]
r = 4
for label, val in readme_lines:
    if val is True:
        ws.cell(row=r, column=1, value=label).font = Font(name=FONT, bold=True, size=12, color="1F4E5F")
        r += 1
        continue
    c1 = ws.cell(row=r, column=1, value=label)
    c1.font = Font(name=FONT, bold=True, size=10)
    c1.alignment = WRAP
    c2 = ws.cell(row=r, column=2, value=val)
    c2.font = BODY_FONT
    c2.alignment = WRAP
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
    r += 1
set_widths(ws, [22] + [12]*7)
ws.column_dimensions["B"].width = 90

# ---------------------------------------------------------------------------
# Causes_Taxonomy
# ---------------------------------------------------------------------------
ws = wb.create_sheet("Causes_Taxonomy")
causes_headers = ["Cause_Code", "Category", "Subcategory", "Combined_Label", "Description", "Prevention_Training_Point"]
causes_rows_raw = [
    ("ELEC-SC", "Electrical", "Short Circuit", "Faulty, aged, or damaged wiring insulation causing arcing/ignition.", "Periodic electrical audits and thermal imaging; replace wiring past rated life; avoid daisy-chained extensions."),
    ("ELEC-OL", "Electrical", "Overloaded Circuit/Panel", "Too many high-draw devices on one circuit or panel exceeding rated capacity.", "Load-balance circuits; dedicated circuits for high-draw equipment (autoclaves, imaging, AC); enforce no unauthorized extension boards."),
    ("ELEC-SG", "Electrical", "Faulty Switchgear/Distribution Board", "Poor-quality, aged, or water-damaged MCBs/DBs/panels.", "Scheduled switchgear replacement/testing; IP-rated enclosures in damp areas; qualified-electrician sign-off after any modification."),
    ("OXY-CYL", "Oxygen System", "Cylinder Leak/Explosion", "Leaking valves, physical damage, or cylinders stored near heat/flame/friction.", "Secure cylinder storage away from ignition sources and direct sunlight; valve inspection before use; staff training on safe handling and transport."),
    ("OXY-ENR", "Oxygen System", "Oxygen-Enriched Atmosphere Combustion", "Elevated O2 concentration in ICU/ward air dramatically accelerates fire spread and intensity once ignited.", "Ventilation design to prevent O2 pooling; no naked flame/smoking/static-generating material near O2 outlets; O2 flow shut-off drills."),
    ("OXY-MGPS", "Oxygen System", "Piped Medical Gas Pipeline Failure", "Leaks, damaged joints, or valve failure in the central medical gas pipeline system.", "Scheduled MGPS integrity testing and leak detection; trained biomedical staff for maintenance; emergency shut-off valve drills."),
    ("HVAC-AC", "HVAC/Cooling", "AC Unit Overheating/Short Circuit", "Poorly maintained compressors, capacitors, or wiring in split/central AC units.", "Annual AC servicing schedule; thermal cut-off devices; dedicated circuit breakers per unit."),
    ("GEN-DG", "Backup Power", "DG Set/Generator Fire", "Fuel leakage, overheating, or exhaust-related ignition in diesel generator sets.", "Generator room fire suppression and ventilation; fuel storage separated from patient areas; scheduled maintenance."),
    ("GEN-UPS", "Backup Power", "UPS/Battery Failure", "Battery overheating or thermal runaway in UPS/inverter battery banks.", "Dedicated ventilated battery rooms; battery health monitoring; replace batteries per manufacturer life rating."),
    ("HUM-NEG", "Human Factors", "Negligence/Unattended Equipment", "Staff error, unattended heat-generating equipment, or bypassed safety interlocks.", "Staff SOP training; equipment shutdown checklists; supervision of high-risk tasks (e.g. welding/repair work in-facility)."),
    ("HUM-FLM", "Human Factors", "Smoking/Open Flame Violation", "Smoking or open flame near oxygen sources or flammable stores.", "Strict no-smoking enforcement with signage; visitor and staff compliance checks near O2 zones."),
    ("FS-SUPP", "Structural/Fire Safety", "Absent/Non-functional Fire Suppression", "No sprinklers, expired extinguishers, or non-functional fire alarm/detection systems.", "Mandatory fire-safety audits and NOC renewal; monthly extinguisher and alarm checks; sprinkler coverage in high-risk zones (ICU, O2 store, electrical room)."),
    ("FS-EXIT", "Structural/Fire Safety", "Blocked Fire Exits/Poor Evacuation Design", "Overcrowded corridors, locked exits, or inadequate egress routes for bed-bound patients.", "Regular evacuation drills including bed-bound patient transfer; keep egress routes clear; horizontal-evacuation zoning for ICUs."),
    ("KIT-FIRE", "Kitchen/Boiler", "Kitchen Fire", "LPG leaks, unattended cooking equipment, or boiler malfunction in hospital kitchens.", "LPG leak detectors and shut-off valves; kitchen fire suppression hoods; separate kitchen from patient-care blocks where possible."),
    ("ARSON", "Arson/Deliberate", "Arson", "Deliberate ignition.", "Security monitoring and CCTV coverage of high-risk areas; access control."),
    ("MEDEQ-MAL", "Medical Equipment", "Equipment Malfunction/Overheating", "Internal fault, overheating, or electrical failure within a medical device itself (e.g. infant incubator/radiant warmer, ventilator, oxygen concentrator) rather than the building's fixed electrical system.", "Preventive-maintenance schedules per manufacturer service intervals; biomedical engineering inspection and fault-code logging; immediate quarantine of any device showing fault indicators, unusual heat, or smell; redundant/backup critical-care equipment (especially neonatal) so a faulty unit can be swapped out immediately rather than run to failure."),
    ("UNK", "Unknown/Under Investigation", "Unknown", "Cause not yet determined, investigation ongoing, or competing hypotheses not yet resolved.", "Update this record once official investigation findings are published."),
]
causes_rows = [(code, cat, sub, f"{cat} - {sub}", desc, prev) for code, cat, sub, desc, prev in causes_rows_raw]
last = write_table(ws, causes_headers, causes_rows, widths=[12, 20, 30, 40, 55, 55])
CAUSES_LAST_ROW = last

# ---------------------------------------------------------------------------
# Products_Taxonomy
# ---------------------------------------------------------------------------
ws = wb.create_sheet("Products_Taxonomy")
prod_headers = ["Product_Code", "Product_Equipment_Category", "Examples", "Common_Failure_Modes", "Relevant_Standards"]
prod_rows = [
    ("PR-ELEC", "Electrical Distribution", "MCBs, distribution boards, wiring/cabling, sockets", "Insulation breakdown, loose connections, overload, arcing, corrosion at joints", "IS 732, IS 8623, National Building Code (NBC) 2016 Part 4 (Fire and Life Safety)"),
    ("PR-OXYCYL", "Oxygen Cylinders", "Portable O2 cylinders, cylinder trolleys/manifolds", "Valve leakage, physical damage/dents, overpressure, corrosion, incompatible regulators", "IS 3224, IS 7285, Gas Cylinder Rules 2016 (PESO)"),
    ("PR-MGPS", "Medical Gas Pipeline System (MGPS)", "Piped oxygen lines, manifolds, zone valves", "Pipe-joint leaks, valve failure, material degradation, incorrect installation", "IS 15498; NFPA 99 referenced internationally for MGPS design"),
    ("PR-CONC", "Oxygen Concentrators", "Bedside O2 concentrators", "Overheating, internal short circuit, filter/compressor fire, continuous-duty failure", "IEC 60601-1, ISO 80601-2-69"),
    ("PR-HVAC", "Air Conditioning/HVAC", "Split ACs, central HVAC, compressors, ducting", "Capacitor failure, refrigerant leak, overheated/seized compressor, dust-clogged coils", "IS 1391 (Room Air Conditioners)"),
    ("PR-BACKUP", "Backup Power", "Diesel generators, UPS units, inverter batteries", "Fuel leak, overheating, exhaust ignition, battery thermal runaway", "IS 4722; NBC 2016 Part 4 (fire and life safety provisions for generator/battery rooms)"),
    ("PR-MEDEQ", "Medical Equipment (General)", "Ventilators, monitors, infusion pumps, defibrillators, warmers", "Internal component failure, overheating power supply/adapter, battery fault", "IEC 60601 series"),
    ("PR-KITCH", "Kitchen Equipment", "LPG stoves, cylinders, exhaust hoods, boilers", "Gas leak, grease fire, exhaust duct fire, unattended equipment", "IS 15656 (LPG installations - code of practice)"),
    ("PR-BUILD", "Building Materials/Furnishings", "False ceilings, curtains, mattresses, PVC cable insulation, foam upholstery", "Rapid flame spread, dense toxic smoke generation, dripping molten plastic", "IS 1642 (Fire Safety of Buildings), NBC 2016 Part 4"),
    ("PR-FIRESYS", "Firefighting/Detection Systems", "Smoke detectors, sprinklers, fire extinguishers, hydrants, fire alarm panels", "Non-functional/expired equipment, poor maintenance, inadequate coverage", "IS 2189 (Automatic Fire Detection), IS 15105, NBC 2016 Part 4"),
]
last = write_table(ws, prod_headers, prod_rows, widths=[12, 28, 40, 50, 50])
PROD_LAST_ROW = last

# ---------------------------------------------------------------------------
# Sources
# ---------------------------------------------------------------------------
ws = wb.create_sheet("Sources")
src_headers = ["Source_ID", "Linked_Incident_IDs", "Source_Type", "Publication_Agency", "Title", "URL",
               "Publication_Date", "Retrieved_Date", "Reliability_Notes"]
src_rows = [
    ("S001", "HF-2026-001", "News - International Wire", "Al Jazeera",
     "Fire at India hospital intensive care unit kills 10 people",
     "https://www.aljazeera.com/news/2026/3/16/fire-at-india-hospital-intensive-care-unit-kills-10-people",
     "2026-03-16", "2026-09-13", "High - international outlet, consistent with other wire coverage of the same event."),
    ("S002", "HF-2026-002", "News - National", "The Week",
     "Bihar hospital blaze: Probe underway after massive fire at ICU ward leaves at least three dead",
     "https://www.theweek.in/news/india/2026/06/04/bihar-muzaffarpur-hospital-fire-icu.html",
     "2026-06-04", "2026-09-13", "Medium-High - national outlet; reports 3 confirmed dead, toll 'could rise'."),
    ("S003", "HF-2026-002", "News - International Wire", "Xinhua",
     "5 killed, 20 injured in India's hospital fire",
     "https://english.news.cn/asiapacific/20260604/1cc53d3afcfa44879185883c4c0ace4a/c.html",
     "2026-06-04", "2026-09-13", "CONFLICTS with S002 on death toll (5 vs 3) - likely reflects toll rising same-day or wire/state-source discrepancy. Flagged in Incidents.Notes; verify against official record when available."),
    ("S004", "HF-2026-003", "Medical Trade Press", "Medical Dialogues",
     "Short circuit triggers fire at Kanpur's LLR Hospital, sparks panic",
     "https://medicaldialogues.in/news/health/hospital-diagnostics/short-circuit-triggers-fire-at-kanpurs-llr-hospital-sparks-panic-175131",
     "2026-07-16", "2026-09-13", "Medium - single trade-press source, quotes hospital Chief Superintendent directly; no independent corroboration logged yet."),
    ("S005", "HF-2026-004", "News - Regional", "Dainik Jagran English (dainikjagranmpcg)",
     "Rewa National Hospital fire after short circuit",
     "https://english.dainikjagranmpcg.com/states/madhya-pradesh/vindhya-rewa/rewa-national-hospital-fire-after-short-circuit/article-18874",
     "2026-05-20", "2026-09-13", "Medium - regional digital outlet; full article blocked automated retrieval (403), details drawn from search-index summary only - re-verify original article before formal use."),
    ("S006", "HF-2026-005", "News - Regional", "Newskarnataka",
     "Fire at Bengaluru hospital, 21 patients escape safely",
     "https://newskarnataka.com/bengaluru/fire-at-bengaluru-hospital-21-patients-escape-safely/27052026/",
     "2026-05-27", "2026-09-13", "High - fetched in full; detailed, internally consistent, names police investigating unit."),
    ("S007", "HF-2026-005", "News - National", "Deccan Herald",
     "21 patients evacuated after fire breaks out in Bengaluru hospital",
     "https://www.deccanherald.com/india/karnataka/bengaluru/21-patients-evacuated-after-fire-breaks-out-in-bengaluru-hospital-4017735",
     "2026-05-27", "2026-09-13", "Medium-High - corroborates S006 (21 evacuated, no injuries) via search index; full article blocked automated retrieval (403)."),
    ("S008", "HF-2026-006", "News - Regional", "Dainik Jagran English (dainikjagranmpcg)",
     "Rewa hospital fire: SNCU ward short circuit, newborns safe",
     "https://english.dainikjagranmpcg.com/states/madhya-pradesh/vindhya-rewa/rewa-hospital-fire-sncu-ward-short-circuit-newborns-safe/article-19400",
     "2026-05-29", "2026-09-13", "Medium - regional digital outlet; full article blocked automated retrieval (403), details drawn from search-index summary only - re-verify original article before formal use."),
    ("S009", "HF-2026-007", "News - Regional (Hindi)", "Prabhat Khabar (Bihar)",
     "सदर अस्पताल में शॉर्ट सर्किट से निकली आग, धुएं से मची अफरातफरी (Short circuit sparks fire, smoke causes panic at Sadar Hospital)",
     "https://www.prabhatkhabar.com/state/bihar/kaimur/fire-breaks-out-sadar-hospital-short-circuit",
     "2026-09-02", "2026-09-13", "Medium - regional Hindi daily, fetched and translated in full; no independent corroborating source found yet."),
    ("S010", "HF-2026-008", "News - Regional (Hindi)", "Nav Bharat Live (Madhya Pradesh)",
     "दतिया जिला अस्पताल में आग से बड़ा हादसा टला (Fire averted a major mishap at Datia District Hospital)",
     "https://navbharatlive.com/madhya-pradesh/datia/datia-district-hospital-old-maternity-ward-fire-short-circuit",
     "2026-08-28", "2026-09-13", "Medium - regional Hindi outlet, fetched and translated in full; no independent corroborating source found yet."),
    ("S011", "HF-2026-009", "News - Regional (Hindi)", "Nav Bharat Live (Maharashtra)",
     "आग लगने के बाद ही क्यों जागता है सिस्टम? NICU से ICU तक खतरा (Why does the system wake up only after a fire? NICU fire at Asian Hospital, Chhatrapati Sambhajinagar)",
     "https://navbharatlive.com/maharashtra/chhatrapati-sambhajinagar/asian-hospital-nicu-fire-nine-newborns-among-58-patients-safe-chhatrapati-sambhajinagar-auk01",
     "2026-09-04", "2026-09-13", "Medium-High - regional Hindi outlet, fetched and translated in full; explicitly ties the incident to the Amravati fire (S017/S018) as a systemic pattern."),
    ("S012", "HF-2026-009", "News - National (Hindi)", "India TV Hindi",
     "हॉस्पिटल के NICU वार्ड में आग, समय रहते बचाए गए 9 नवजात (Fire in hospital NICU ward, 9 newborns saved in time)",
     "https://www.indiatv.in/maharashtra/chhatrapati-sambhajinagar-asian-hospital-nicu-fire-short-circuit-9-newborns-shifted-2026-09-04-1241226",
     "2026-09-04", "2026-09-13", "Medium-High - national Hindi broadcaster corroborating S011 on patient count and zero-injury outcome (via search index)."),
    ("S013", "HF-2026-010", "News - Regional (Hindi)", "Haribhoomi (Chhattisgarh)",
     "CIMS में शॉर्ट सर्किट से भड़की आग: ICU समेत कई वार्डों की बिजली गुल (Fire from short circuit at CIMS: power cut to ICU and several wards)",
     "https://www.haribhoomi.com/state-local/chhattishgarh/news/bilaspur-cims-hospital-fire-short-circuit-security-arrangement-fire-extinguisher-111459",
     "2026-09-04", "2026-09-13", "Medium-High - regional Hindi outlet, fetched and translated in full; names outdated wiring as a contributing factor."),
    ("S014", "HF-2026-010", "News - Regional (Hindi)", "NPG News (Chhattisgarh)",
     "CG News: अस्पताल में शॉर्ट सर्किट से भड़की भीषण आग (Massive fire from short circuit at hospital)",
     "https://npg.news/chhattisgarh/cg-bilaspur-cims-me-lagi-aag-latest-cg-news-hindi-npg-04-09-2026-1319264",
     "2026-09-04", "2026-09-13", "Medium - regional Hindi outlet corroborating S013 (via search index); notes doctors treated patients by phone-torch light during the outage."),
    ("S015", "HF-2026-011", "News - Regional (Hindi)", "Amar Ujala (Bareilly)",
     "बरेली के नवोदय अस्पताल की चौथी मंजिल पर लगी आग, मरीजों को सुरक्षित निकाला गया (Fire on 4th floor of Navodaya Hospital, Bareilly; patients evacuated safely)",
     "https://www.amarujala.com/uttar-pradesh/bareilly/fire-breaks-out-on-the-fourth-floor-of-navodaya-hospital-in-bareilly-patients-safely-evacuated-2026-08-29",
     "2026-08-29", "2026-09-13", "Medium-High - major regional Hindi daily, fetched and translated in full."),
    ("S016", "HF-2026-011", "News - Regional (Hindi)", "Subkuz",
     "बरेली अस्पताल मे आग, मरीज सुरक्षित बाहर निकाले (Fire at Bareilly hospital, patients evacuated safely)",
     "https://subkuz.com/news/hindi/details/bareilly-navodaya-hospital-fire-five-storey-building-patients-evacuated-safely/213299",
     "2026-08-29", "2026-09-13", "Medium - corroborates S015 on evacuation and zero-casualty outcome (via search index)."),
    ("S017", "HF-2026-012", "News - National", "The Quint",
     "Three newborns die in Amravati hospital NICU fire, probe underway",
     "https://www.thequint.com/news/breaking-news/maharashtra-newborns-die-amravati-hospital-fire",
     "2026-08-24", "2026-09-13", "High - national outlet, fetched in full; names DCP and Revenue Minister directing the inquiry."),
    ("S018", "HF-2026-012", "News - National", "The Tribune",
     "Amravati hospital fire: Braved thick smoke, darkness to rescue charred babies: Nurses",
     "https://www.tribuneindia.com/news/india/amravati-hospital-fire-braved-thick-smoke-darkness-to-rescue-charred-babies-nurses/",
     "2026-08-24", "2026-09-13", "High - national outlet corroborating S017, with first-hand nursing staff accounts of the rescue (via search index)."),
    ("S019", "HF-2026-013", "News - Regional (English)", "OdishaTV",
     "Fire breaks out at MKCG Medical College and Hospital in Berhampur, no casualties reported",
     "https://odishatv.in/odisha/fire-breaks-out-at-mkcg-medical-college-and-hospital-in-berhampur-no-casualties-reported-12536265",
     "2026-09-15", "2026-09-18", "High - major Odisha outlet, fetched in full; timestamped 15 Sep 2026 18:05 IST and says 'Tuesday', which fixes the incident date as 2026-09-15."),
    ("S020", "HF-2026-013", "Medical Trade Press", "Medical Dialogues",
     "Short circuit in AC triggers fire at MKCG Medical College Hospital in Berhampur",
     "https://medicaldialogues.in/news/health/hospital-diagnostics/short-circuit-in-ac-triggers-fire-at-mkcg-medical-college-hospital-in-berhampur-179344",
     "2026-09-16", "2026-09-18", "Medium-High - fetched in full; source of the expired-extinguisher detail (refilled 2025-08-01, validity expired 2026-07-31) and the fire officer declining to answer. Publication date is Sep 16; the fire itself was Tuesday Sep 15 per S019."),
    ("S021", "HF-2026-013", "News - Regional (Odia)", "ETV Bharat Odia",
     "ରାଜ୍ୟର ଟପ୍ ମେଡିକାଲରେ ନାମକୁ ମାତ୍ର ଅଗ୍ନି ସୁରକ୍ଷା ବ୍ୟବସ୍ଥା (Only nominal fire-safety arrangements at the state's top medical colleges)",
     "https://www.etvbharat.com/or/state/explainer-mkcg-medical-college-scb-medical-college-hospital-fire-safety-smoke-control-fire-drills-ors26091606070",
     "2026-09-16", "2026-09-18", "Medium-High - explainer fetched and translated; documents lapsed fire-safety certificates, expired extinguishers, untrained staff and the 2026 MKCG fire history. Publication date inferred from the URL id."),
    ("S022", "HF-2026-013", "News - Regional (English)", "Argus English",
     "Repeated fires at MKCG Medical College raise safety concerns in Odisha (third fire in 23 days)",
     "https://argusenglish.in/odisha/mkcg-medical-college-and-hospital-sees-third-fire-in-23-days",
     "2026-09-16", "2026-09-18", "Medium - fetched in full; useful for the repeat-fire cluster, but its table dates the Ob/Gyn OT fire 'September 10' and an earlier fire 'August 23', conflicting with S019/S021 (Sep 15; Aug 19). Treated as date errors."),
    ("S023", "HF-2026-009", "News - Regional (Marathi)", "ETV Bharat Marathi",
     "रुग्णालयातील आगीच्या घटनेत बचावले, मात्र नियतीनं हेरलंच! जुळ्या बालकांचा उपचारादरम्यान मृत्यू (Survived the hospital fire, but twin babies die during treatment)",
     "https://www.etvbharat.com/mr/state/twin-babies-who-survived-the-fire-at-asian-hospital-in-chhatrapati-sambhajinagar-died-during-treatment-mhs26090701549",
     "2026-09-07", "2026-09-18", "Medium-High - fetched and translated; quotes the hospital director and a parent; clinicians attribute the deaths to extreme prematurity, not the fire."),
    ("S024", "HF-2026-002", "News - National (Hindi)", "ETV Bharat Hindi",
     "मुजफ्फरपुर अस्पताल अग्निकांड में बड़ी कार्रवाई, डॉक्टर समेत तीन गिरफ्तार, मरने वालों का आंकड़ा पहुंचा 6 (Three arrested including a doctor; death toll reaches 6)",
     "https://www.etvbharat.com/hi/bharat/muzaffarpur-hospital-fire-action-three-arrested-including-doctor-brn26060503709",
     "2026-06-05", "2026-09-18", "Medium-High - fetched and translated; gives 6 dead, 17 injured, names the three arrested. Supersedes the 3 (S002) and 5 (S003) same-day tolls."),
    ("S025", "HF-2026-002", "News - National", "ETV Bharat English",
     "Police Arrest 3 In Muzaffarpur Hospital Fire Case; Death Toll Up To 6",
     "https://www.etvbharat.com/en/bharat/police-arrest-3-in-muzaffarpur-hospital-fire-case-death-toll-up-to-6-enn26060504108",
     "2026-06-05", "2026-09-18", "Medium-High - corroborates S024 (toll 6, three arrests, owner absconding; fire dept preliminary view that inadequate ICU fire safety let the fire spread) via search index; not fetched separately."),
    ("S026", "HF-2026-002", "News - National (Hindi)", "Patrika",
     "मुजफ्फरपुर अस्पताल अग्निकांड में बड़ा खुलासा, ICU में लापरवाही से मरीजों की मौत, फायर सिस्टम फेल (Big revelation: ICU negligence, fire system failed)",
     "https://www.patrika.com/muzaffarpur-news/muzaffarpur-prasad-hospital-fire-icu-negligence-3-dead-overcrowding-fire-safety-failure-20638299",
     "2026-06-05", "2026-09-18", "Low - not fetched; search-index title cites 10 deaths while the URL slug says '3 dead' and every other source says 6. Retained for transparency only; the ICU 13-bed/15-patient overcrowding claim in HF-2026-002 comes from search summaries of Hindi press coverage and is unverified."),
    ("S027", "HF-2026-014", "News - National", "ETV Bharat English",
     "Telangana RIMS Hospital Fire: Toll Rises To 5 As Two More Infants Succumb To Burn Injuries",
     "https://www.etvbharat.com/en/state/telangana-adilabad-rims-hospital-fire-toll-rises-to-five-as-two-more-children-succumb-to-burn-injuries-enn26092302979",
     "2026-09-23", "2026-09-28", "Medium-High - fetched in full; gives the toll progression (1, then 3, then 5) and smoke inhalation as cause of death. Its own text dates the fire 'September 22-23', but the other sources put it on the night of Monday Sep 21; the headline says 'burn injuries' while the body says smoke inhalation."),
    ("S028", "HF-2026-014", "News - Regional (English)", "The South First",
     "Final 5 minutes before smoke killed the newborn: The fiery night Adilabad RIMS lost three babies",
     "https://thesouthfirst.com/health/final-5-minutes-before-smoke-killed-the-newborn-the-fiery-night-adilabad-rims-lost-three-babies/",
     "2026-09-22", "2026-09-28", "Medium-High - fetched in full; the only report with a minute-level account; source for the locked glass doors (security broke them) and the electrical, fire-safety and structural audits ordered. Written when the toll was 3."),
    ("S029", "HF-2026-014", "Medical Trade Press", "Medical Dialogues",
     "2 infants die after fire breaks out at RIMS Adilabad neonatal unit",
     "https://medicaldialogues.in/news/health/hospital-diagnostics/2-infants-die-after-fire-breaks-out-at-rims-adilabad-neonatal-unit-179726",
     "2026-09-22", "2026-09-28", "Medium - fetched in full; earliest tally (2 dead, 28 infants) and the fire official's statement that the fire 'appears to have originated from the AC due to a short-circuit'. Superseded by later tolls."),
    ("S030", "HF-2026-014", "News - Regional (English)", "Telangana Today",
     "Six-member committee begins probe into RIMS-Adilabad tragedy",
     "https://telanganatoday.com/six-member-committee-begins-probe-into-rims-adilabad-tragedy",
     "2026-09-24", "2026-09-28", "Medium-High - fetched in full; committee composition and mandate, toll of 5 (one immediate, four during treatment after inhaling toxic gases), Rs 5 lakh compensation. Gives no AC specifics."),
    ("S031", "HF-2026-014", "News - Regional (English)", "Telangana Today",
     "Poor management of electrical system, ACs blamed for RIMS-Adilabad tragedy",
     "https://telanganatoday.com/poor-management-of-electrical-system-acs-blamed-for-rims-adilabad-tragedy",
     "2026-09-23", "2026-09-28", "Low-Medium - fetched in full, but every maintenance claim (130+ ACs bought in 2016, no regular electricians, missed quarterly inspections, delayed repairs, sub-standard refrigerant, compressor overload) is attributed to unnamed officials or an unnamed expert, not to the inquiry. Treat as allegations to test against the committee report."),
    ("S032", "HF-2026-014", "News - Regional (English)", "Deccan Chronicle",
     "Revenue officials sealed the SNCU ward at RIMS, Adilabad",
     "https://www.deccanchronicle.com/southern-states/telangana/revenue-officials-sealed-the-sncu-ward-at-rims-adilabad-1990078",
     "2026-09-24", "2026-09-28", "Medium - fetched in full; ward sealed to prevent evidence tampering, committee to report within 10 days."),
    ("S033", "HF-2026-014", "News - Regional (English)", "Siasat",
     "Telangana: Death toll rises to five in Adilabad RIMS fire",
     "https://www.siasat.com/telangana-death-toll-rises-to-five-in-adilabad-rims-fire-3546924/",
     "2026-09-23", "2026-09-28", "Medium - independent corroboration of the toll of 5, via search index; not fetched."),
    ("S034", "HF-2026-015", "News - Regional (Hindi)", "Amar Ujala (Darbhanga)",
     "बिहार: डीएमसीएच के एमसीएच वार्ड में शॉर्ट सर्किट से लगी आग, मरीजों में मची अफरा-तफरी; सभी सुरक्षित शिफ्ट (Short circuit fire in DMCH MCH ward; all patients shifted safely)",
     "https://www.amarujala.com/bihar/darbhanga/bihar-news-short-circuit-caused-fire-mch-ward-dmch-panic-patients-shifted-safety-darbhanga-news-c-1-1-noi1239-4766840-2026-09-26",
     "2026-09-26", "2026-09-28", "Medium-High - major regional Hindi daily, fetched and translated in full."),
    ("S035", "HF-2026-015", "News - Regional (Hindi)", "Lalluram",
     "DMCH के MCH वार्ड में देर रात शॉर्ट सर्किट से लगी आग, मरीजों में मची अफरा-तफरी (Late-night short circuit fire in DMCH MCH ward)",
     "https://lalluram.com/dmch-mch-ward-fire-incident-darbhanga/",
     "2026-09-26", "2026-09-28", "Medium - fetched and translated; corroborates S034 (electrical panel short circuit, pregnant women and newborns evacuated, no casualties)."),
    ("S036", "HF-2026-015", "News - Regional (Hindi)", "Prabhat Khabar (Bihar)",
     "DMCH MCH भवन में लगी आग, 2 दर्जन से अधिक मां-बच्चे सुरक्षित बचे (Fire in DMCH MCH building; over two dozen mothers and babies rescued)",
     "https://www.prabhatkhabar.com/state/bihar/darbhanga/dmch-mother-child-hospital-fire",
     "2026-09-26", "2026-09-28", "Medium - not fetched (page returned HTTP 410); headline seen via search index and is the only source for the 'over two dozen' rescued figure."),
    ("S037", "HF-2026-014", "News - Regional (English)", "Telangana Today",
     "Three newborns die as ACs explode at SNCU in Adilabad",
     "https://telanganatoday.com/three-newborns-die-as-acs-explode-at-sncu-in-adilabad",
     "2026-09-22", "2026-09-28", "Low-Medium - not fetched; seen via search-index summary only. Source, with S028, for the doors locking after the blast and the unverified allegation that duty doctors and nurses fled; also the source for 'ACs' (plural) exploding."),
]
last = write_table(ws, src_headers, src_rows, widths=[10, 16, 22, 22, 45, 45, 14, 14, 55])
SRC_LAST_ROW = last

# ---------------------------------------------------------------------------
# Incidents
# ---------------------------------------------------------------------------
ws = wb.create_sheet("Incidents")
inc_headers = [
    "Incident_ID", "Date_of_Incident", "State", "City_District", "Hospital_Name", "Hospital_Type",
    "Ward_Location_of_Origin", "Fatalities", "Injuries", "Patients_Evacuated", "Severity_Category",
    "Root_Cause_Category", "Root_Cause_Detail", "Product_Equipment_Involved", "Failure_Mode",
    "Oxygen_Related", "Fire_Safety_NOC_Status", "Fire_Safety_Violations_Noted",
    "Investigation_Status", "Investigation_Findings_Legal_Action", "Data_Confidence",
    "Source_IDs", "Date_Logged", "Logged_By", "Notes",
]
col = {h: get_column_letter(i + 1) for i, h in enumerate(inc_headers)}
inc_rows = [
    ("HF-2026-001", "2026-03-16", "Odisha", "Cuttack", "SCB Medical College and Hospital", "Government",
     "Trauma care unit, first floor (ICU)", 10, 16, 13, "Fatal",
     "Electrical - Short Circuit", "Short circuit attributed to poorly maintained wiring in the trauma ICU; 23 patients were under intensive care when the fire broke out.",
     "Electrical Distribution", "Short Circuit / Arcing",
     "Unclear", "Unknown", "Not specified in reporting to date.",
     "Ongoing", "Odisha CM ordered a judicial inquiry; PM Modi announced ex-gratia compensation of $2,160 per affected family.",
     "Confirmed - Multiple News Sources", "S001", "2026-09-13", "gdpravin@yahoo.com",
     "10 deaths occurred during evacuation, not from direct burns. Staff and security personnel reportedly risked injury rescuing patients. Injuries = 11 staff burns + 5 critically injured patients per source."),
    ("HF-2026-002", "2026-06-04", "Bihar", "Muzaffarpur (Brahmpura)", "Prasad Hospital", "Private",
     "Fifth floor ICU (fire also threatened adjoining CCU)", 6, 17, None, "Fatal",
     "Electrical - Short Circuit", "Preliminary reports point to a short circuit at ~3:55 AM in the fifth-floor ICU of a private multi-storey hospital.",
     "Electrical Distribution", "Short Circuit (preliminary/unconfirmed)",
     "Unclear", "Unknown", "Fire department's preliminary view: the fire spread quickly because of inadequate fire-safety measures in the ICU. Search summaries of Hindi press coverage also report the ICU (13 beds) held 15 patients - overcrowding, unverified against the original articles. Rapid spread hampered evacuation.",
     "Ongoing", "Three arrested on 2026-06-05 for failing to perform their responsibilities: ICU in-charge Dr Pankaj, administrative manager Ramkumar and maintenance manager Ajit Kumar. Owner Dr Upendra Prasad and family reported absconding at that time. District Magistrate ordered an investigation; Hindi press also reports an FIR ordered against the owner and manager and a human-rights-commission petition seeking a retired-judge inquiry (headline-level, unverified).",
     "Confirmed - Multiple News Sources", "S002, S003, S024, S025, S026", "2026-09-13", "gdpravin@yahoo.com",
     "Updated 2026-09-18: the earlier 3-vs-5 discrepancy was the toll rising during the day - ETV Bharat (S024/S025) reports 6 dead and 17 injured by 2026-06-05, so Fatalities/Injuries now use those figures (Xinhua S003 said 20 injured). Check for further toll rises. One Hindi outlet's headline (S026) cites 10 deaths; it contradicts every other source and is treated as an outlier. The three arrests make this the only incident so far with criminal accountability for hospital staff, which is relevant to the training and governance angle."),
    ("HF-2026-003", "2026-07-16", "Uttar Pradesh", "Kanpur", "Lala Lajpat Rai (LLR) Hospital", "Government",
     "Electric panel near ENT department", 0, 0, 100, "Near Miss (No Injury)",
     "Electrical - Short Circuit", "Short circuit at the electric panel near the ENT department, per the Hospital Chief Superintendent; heavy smoke suspended ultrasound/X-ray services temporarily.",
     "Electrical Distribution", "Short Circuit",
     "No", "Unknown", "Not specified in reporting to date.",
     "Ongoing", "Hospital administration reviewing safety measures to prevent recurrence; no casualties or major damage reported.",
     "Single Source - Unverified", "S004", "2026-09-13", "gdpravin@yahoo.com",
     "No deaths/injuries reported. Fire was contained quickly with hospital-staff-operated fire extinguishers - a positive-outcome case worth including in training material as an example of effective initial response."),
    ("HF-2026-004", "2026-05-20", "Madhya Pradesh", "Rewa", "National Hospital", "Unknown",
     "Corridor near outpatient (OPD) block", 0, "Minor - unspecified count (smoke inhalation, treated & released)", None, "Non-Fatal Injury",
     "Electrical - Short Circuit", "Short circuit in a hospital wing sparked a fire in a corridor near the OPD block at approx. 2:10 pm; patients, attendants and staff rushed out of the building.",
     "Electrical Distribution", "Short Circuit",
     "No", "Unknown", "Not specified in reporting to date.",
     "Ongoing", "Hospital stated a safety review and electrical audit would follow.",
     "Single Source - Unverified", "S005", "2026-09-13", "gdpravin@yahoo.com",
     "Near-miss/minor-injury case: no deaths; a few people treated on-site for mild smoke-inhalation and released. Source site blocked automated full-article retrieval (403) - details drawn from a search-index summary only; re-verify against the original article before formal use."),
    ("HF-2026-005", "2026-05-27", "Karnataka", "Bengaluru (Singapura, M.S. Palya Road)", "Aveksha Hospital", "Private",
     "Basement electrical duct; smoke/fire spread to 2nd-floor ICU", 0, 0, 21, "Near Miss (No Injury)",
     "Electrical - Short Circuit", "Fire suspected to have originated from an electrical cable in a basement duct at approx. 3:15 am; ICU (2nd floor) and 1st floor evacuated.",
     "Electrical Distribution", "Short Circuit (suspected, cable fault in basement duct)",
     "Unclear", "Unknown", "Not specified in reporting to date.",
     "Ongoing", "Vidyaranyapura police investigating exact cause; final damage assessment pending. ICU area sustained smoke and fire damage.",
     "Confirmed - Multiple News Sources", "S006, S007", "2026-09-13", "gdpravin@yahoo.com",
     "Staff evacuated all 21 patients (14 from ICU, 7 from first floor) before fire services arrived - a positive-outcome near-miss useful for evacuation-drill training. One less-detailed source reported '14 patients evacuated' (the ICU subset only); reconciled here as 21 total per the fuller accounts."),
    ("HF-2026-006", "2026-05-29", "Madhya Pradesh", "Rewa", "Gandhi Memorial Hospital", "Unknown",
     "SNCU (Special Newborn Care Unit) ward", 0, 0, None, "Near Miss (No Injury)",
     "Electrical - Short Circuit", "Sparks from an electrical short circuit spread quickly within the SNCU ward at approx. 11:30 pm; newborns reported safe.",
     "Electrical Distribution", "Short Circuit",
     "Unclear", "Unknown", "Not specified in reporting to date.",
     "Unknown", "Not specified in reporting to date.",
     "Single Source - Unverified", "S008", "2026-09-13", "gdpravin@yahoo.com",
     "High-severity potential given the SNCU (neonatal) setting, occurring just 9 days after HF-2026-004 in the same city - flags Rewa, MP as worth watching for a localized pattern of aging hospital electrical infrastructure once more sources are found. Source site blocked automated full-article retrieval (403) - details drawn from a search-index summary only; re-verify before formal use."),
    ("HF-2026-007", "2026-09-02", "Bihar", "Bhabhua, Kaimur", "Sadar Hospital, Bhabhua", "Government",
     "Electrical wiring near the water tank, outside the Emergency ward", 0, 0, None, "Near Miss (No Injury)",
     "Electrical - Short Circuit", "Short circuit in electrical wires near the water tank outside the emergency ward caused sparks and smoke around noon; no injuries reported.",
     "Electrical Distribution", "Short Circuit",
     "No", "Unknown", "Article raises concern about aging electrical infrastructure in a high-traffic emergency-ward area; no formal violation cited.",
     "Unknown", "No formal investigation reported; hospital management urged to assess wiring/connection conditions.",
     "Single Source - Unverified", "S009", "2026-09-13", "gdpravin@yahoo.com",
     "Sourced from Hindi-language regional press (Prabhat Khabar, Bihar edition) as part of a deliberate effort to capture near-miss incidents that go unreported in national English media."),
    ("HF-2026-008", "2026-08-28", "Madhya Pradesh", "Datia", "Datia District Hospital", "Government",
     "Old maternity ward, first floor - storage room behind lab/dialysis section", 0, 0, 0, "Near Miss (No Injury)",
     "Electrical - Short Circuit", "Short circuit (preliminary) in a storage room of the old maternity ward destroyed stored mattresses/linens; the ward had no admitted patients at the time.",
     "Electrical Distribution", "Short Circuit (preliminary)",
     "No", "Unknown", "Not specified in reporting to date.",
     "Ongoing", "Hospital administration investigating cause; Resident Medical Officer confirmed no loss of life.",
     "Single Source - Unverified", "S010", "2026-09-13", "gdpravin@yahoo.com",
     "Zero-casualty outcome here likely reflects the ward being unoccupied at the time rather than fast detection/response - a reminder not to over-credit every zero-injury near miss as a 'good response' case; some are simply lucky on timing."),
    ("HF-2026-009", "2026-09-04", "Maharashtra", "Chhatrapati Sambhajinagar (Akashvani Chowk)", "Asian Hospital (Asian City Care Hospital)", "Private",
     "NICU, first floor - AC unit", 0, 0, 58, "Near Miss (No Injury)",
     "HVAC/Cooling - AC Unit Overheating/Short Circuit", "Suspected AC-unit malfunction/short circuit in the NICU around 2-3 am produced smoke that staff identified immediately, triggering evacuation of all 58 patients including 9 newborns; fire brought under control within 10-15 minutes.",
     "Air Conditioning/HVAC", "Overheating / Short Circuit (AC unit)",
     "Unclear", "Unknown", "Coverage explicitly frames this as part of a pattern of hospital fire-safety systems being scrutinized only reactively, referencing the fatal Amravati NICU fire (HF-2026-012) 11 days earlier.",
     "Unknown", "Hospital director Dr Shoeb Hashmi confirmed a technical malfunction in an AC unit (S023); no formal inquiry reported.",
     "Confirmed - Multiple News Sources", "S011, S012, S023", "2026-09-13", "gdpravin@yahoo.com",
     "Positive-outcome NICU case: staff recognized smoke from the AC unit and evacuated all 9 newborns plus 49 other patients with zero injuries - a strong contrast with the fatal Amravati NICU fire (HF-2026-012) 11 days earlier; useful as a comparative training case study on NICU response speed. See FA-002. Follow-up added 2026-09-18 (S023): twin newborns born at 27 weeks who were moved to Ghati Hospital died on Sep 6-7; clinicians attributed the deaths to extreme prematurity, not fire or smoke exposure, and the family had no complaints. They are NOT counted as fire fatalities; reclassify to Fatal if an inquiry links the deaths to smoke exposure or the emergency transfer."),
    ("HF-2026-010", "2026-09-04", "Chhattisgarh", "Bilaspur", "Chhattisgarh Institute of Medical Sciences (CIMS)", "Government",
     "Main electrical panel, near kitchen/relatives' waiting room", 0, 0, 0, "Near Miss (No Injury)",
     "Electrical - Faulty Switchgear/Distribution Board", "Short circuit at the main electrical panel near the kitchen/relatives' room around 11:30 am cut power to the ICU, TB ward, cancer OPD and kitchen; staff used fire extinguishers and cut mains as a precaution, treating patients by phone-torch light until power was restored.",
     "Electrical Distribution", "Short Circuit (main panel)",
     "Unclear", "Unknown", "Outdated wiring/electrical infrastructure cited explicitly; Public Works Department directed to replace old electrical systems.",
     "Ongoing", "Electrical system inspection underway; hospital credited prior fire-safety training/mock drills for the orderly response.",
     "Confirmed - Multiple News Sources", "S013, S014", "2026-09-13", "gdpravin@yahoo.com",
     "Good example of prior mock-drill training translating into a calm, orderly response despite ICU/critical wards losing power - supports TR-001 and TR-006. Also a concrete, named instance of 'outdated wiring' as a contributing factor rather than an inferred one."),
    ("HF-2026-011", "2026-08-29", "Uttar Pradesh", "Bareilly (Izatnagar, Mini Bypass)", "Navodaya Hospital", "Private",
     "Fourth floor meeting hall (sofa ignition); floor under construction/renovation", 0, 0, 45, "Near Miss (No Injury)",
     "Unknown/Under Investigation - Unknown", "Fire started on a sofa in a fourth-floor meeting hall around 9:30 am; suspected cause is either a short circuit in an AC unit or a discarded lit bidi left by construction workers in the area - cause not conclusively established. No patients were on the affected floor. The fire brigade needed about 90 minutes and had to access the roof via an adjoining hospital due to limited exit routes at Navodaya itself.",
     "Building Materials/Furnishings", "Flame spread on furnishings (ignition source unconfirmed: AC short circuit vs. discarded smoking material)",
     "No", "Unknown", "Firefighters had to reach the roof through an ADJACENT hospital building due to inadequate direct access/exit routes at Navodaya Hospital itself - a notable structural fire-safety gap independent of the ignition cause.",
     "Ongoing", "Chief Medical Officer directed hospital management to conduct an electrical safety inspection and correct identified deficiencies.",
     "Confirmed - Multiple News Sources", "S015, S016", "2026-09-13", "gdpravin@yahoo.com",
     "Notable for two reasons distinct from the electrical-short-circuit pattern elsewhere in this database: (1) it occurred on a floor under construction/renovation, and (2) firefighting access was hampered by inadequate roof/exit access at the hospital itself. Root cause left as Unknown given two competing, unresolved hypotheses (AC fault vs. smoking material) rather than guessing."),
    ("HF-2026-012", "2026-08-24", "Maharashtra", "Amravati", "Amravati District Women's Hospital", "Government",
     "NICU, third floor", 3, 0, 6, "Fatal",
     "Medical Equipment - Equipment Malfunction/Overheating", "A technical fault in a neonatal thermal unit (warmer)/ventilator triggered a fire in the NICU at approximately 3:15-3:30 am; the fire was confined to one room and extinguished within about 30 minutes, but 3 of the 9 infants present died before they could be saved. The remaining 6 were transferred to a super-specialty hospital.",
     "Medical Equipment (General)", "Overheating / Internal Technical Fault",
     "Unclear", "Unknown", "Investigation examining whether fire safety protocols were followed; not yet detailed in reporting.",
     "Ongoing", "Maharashtra Revenue Minister Chandrashekhar Bawankule ordered a full inquiry; police (DCP Shyam Ghuge) investigating the exact technical cause.",
     "Confirmed - Multiple News Sources", "S017, S018", "2026-09-13", "gdpravin@yahoo.com",
     "Distinct failure pathway from the electrical-wiring/panel pattern dominating the rest of this database: the reported ignition source is internal to a piece of neonatal medical equipment itself (thermal unit/ventilator), not the building's fixed electrical system. Points to device-level preventive maintenance and biomedical-engineering oversight as a mitigation track separate from facility electrical audits. See FA-002 for the NICU-cluster analysis linking this with HF-2026-006 and HF-2026-009."),
    ("HF-2026-013", "2026-09-15", "Odisha", "Berhampur (Ganjam)", "MKCG Medical College and Hospital", "Government",
     "Anaesthesia room attached to the operation theatre, 2nd floor, Obstetrics & Gynaecology dept.", 0, 0, 0, "Near Miss (No Injury)",
     "HVAC/Cooling - AC Unit Overheating/Short Circuit", "Electrical short circuit in an air-conditioning unit around 3 pm destroyed the AC and filled the floor with smoke; no patients were in the operation theatre at the time. Hospital fire-safety staff on the premises responded quickly.",
     "Air Conditioning/HVAC", "Short Circuit (AC unit)",
     "Unclear", "Expired", "An expired fire extinguisher (refilled 2025-08-01, validity expired 2026-07-31) hung at the department entrance and the fire officer declined to answer questions about it. Fire-safety certificates for several blocks (hostels, doctors' quarters, labs, exam halls) reportedly lapsed in 2023 and surgery/paediatrics/administrative blocks have had no certification for three years; the Ob/Gyn block's own status is not specified. Staff present reportedly lacked experience operating extinguishers.",
     "Unknown", "Medical College Superintendent Prof. Sudipa Das promised corrective measures within 48 hours and quarterly electrical audits; no formal inquiry reported.",
     "Confirmed - Multiple News Sources", "S019, S020, S021, S022", "2026-09-18", "gdpravin@yahoo.com",
     "First incident found in the 2026-09-18 weekly pass (window Sep 10-18). Zero-casualty outcome owes much to the theatre being empty, so it should not be read as proof the response worked. Part of a repeat-fire cluster: ETV Bharat Odia (S021) lists four MKCG fires in 2026 - May 28 (TB & Chest ICU wall-fan short circuit), Aug 19 (Biochemistry lab), Sep 4 (nursing-hostel kitchen gas leak) and this one; Argus (S022) counts three in 23 days. The three earlier fires are pre-window and not yet logged as separate incidents. Dates conflict between outlets: Argus says Sep 10 for this fire, but OdishaTV's timestamp and ETV Bharat give Sep 15. NOC status is entered as Expired based on hospital-wide reporting. See FA-003."),
    ("HF-2026-014", "2026-09-21", "Telangana", "Adilabad", "Rajiv Gandhi Institute of Medical Sciences (RIMS)", "Government",
     "Special Newborn Care Unit (SNCU) - window AC unit", 5, "Smoke inhalation - at least 20 infants hospitalised (ETV Bharat); further critical cases reported", 27, "Fatal",
     "HVAC/Cooling - AC Unit Overheating/Short Circuit", "A short circuit in a window air-conditioning unit in the SNCU at about 11:40 pm on Monday Sep 21 caused the AC to burst into flames; plastic components melted and dense smoke filled the ward. Five premature infants died - one during evacuation and four during treatment, all from smoke inhalation. About 27 infants were in the unit and were moved to private and other hospitals.",
     "Air Conditioning/HVAC", "Short Circuit / AC explosion (suspected; plastic components burned)",
     "Unclear", "Unknown", "Doors reportedly locked after the blast and security staff and relatives had to break the glass to reach the babies (S028, S037). Press reports raise concerns about AC and electrical maintenance, fire alarms, fire-fighting equipment and evacuation arrangements, but give no verified specifics. Claims attributed to unnamed officials and an unnamed expert (S031), not the inquiry: 130+ ACs bought in 2016, no regular electricians (three outsourced staff assigned other tasks), missed quarterly inspections, delayed repairs, alleged sub-standard refrigerant and compressor overload. One fire tender attended.",
     "Ongoing", "CM A. Revanth Reddy ordered an inquiry and fire, electrical and structural audits; a six-member committee led by ITDA Utnoor project officer Manda Makarandu (with the district fire officer, electricity SE and TGMIDC engineer) began on Sep 24 with a 10-day deadline. Revenue officials sealed the SNCU to preserve evidence and the state human rights commission took suo motu cognisance. Rs 5 lakh ex gratia per family announced; no arrests or cause finding reported yet.",
     "Confirmed - Multiple News Sources", "S027, S028, S029, S030, S031, S032, S033, S037", "2026-09-28", "gdpravin@yahoo.com",
     "Found in the 2026-09-28 weekly pass (window Sep 18-28) and the deadliest incident since Cuttack (HF-2026-001). Toll progressed 1, 2, 3, then 5 by Sep 23 (S027, S030, S033); early wire reports of 2 dead are superseded, and the toll may still rise because further infants were critical. Infants present vary by outlet from 24 to 28 (27 used, per health officials). ETV Bharat dates the fire 'September 22-23' but the other sources give the night of Sep 21. One report (S037, seen only as a search-index summary) alleges duty staff fled, while S028 describes staff evacuating newborns as smoke spread. This is the first fatal AC fire and the second fatal neonatal-unit fire (after Amravati HF-2026-012) in this database, and the first case where the inquiry may produce AC or maintenance evidence. See FA-002 and FA-003."),
    ("HF-2026-015", "2026-09-26", "Bihar", "Darbhanga", "Darbhanga Medical College and Hospital (DMCH)", "Government",
     "Maternal and Child Health (MCH) ward - electrical panel", 0, 0, "24+ (mothers and newborns)", "Near Miss (No Injury)",
     "Electrical - Faulty Switchgear/Distribution Board", "Short circuit in the electrical panel of the MCH ward at about 1 am filled the ward with thick smoke; pregnant women, mothers and newborns were moved to the lower section of the New Surgical Block. The fire was put out within minutes and the ward's electricity stayed cut while repairs were made.",
     "Electrical Distribution", "Short Circuit (electrical panel)",
     "Unclear", "Unknown", "Not specified in reporting to date; coverage calls for electrical safety audits.",
     "Unknown", "No formal investigation reported; repairs to restore the ward's electrical supply under way.",
     "Confirmed - Multiple News Sources", "S034, S035, S036", "2026-09-28", "gdpravin@yahoo.com",
     "Positive-outcome near miss in a maternity and newborn area: all patients were moved without loss in a night-time event, five days after the fatal Adilabad SNCU fire. Service impact is worth noting - the MCH ward stopped treating patients until power was restored. The 'over two dozen' rescued figure comes only from a Prabhat Khabar headline (S036) that could not be fetched."),
]
last = write_table(ws, inc_headers, inc_rows, widths=[13, 13, 16, 22, 30, 12, 30, 9, 22, 11, 18, 22, 45, 20, 20, 10, 14, 30, 13, 40, 24, 12, 12, 20, 50])
INC_LAST_ROW = last
INC_FIRST_DATA_ROW = 2

# Data validations on Incidents
dv_hospital_type = DataValidation(type="list", formula1='"Government,Private,Trust/Charitable,Nursing Home,Unknown"', allow_blank=True)
dv_severity = DataValidation(type="list", formula1='"Fatal,Non-Fatal Injury,Near Miss (No Injury),Property/Equipment Damage Only,Unknown"', allow_blank=True)
dv_oxy = DataValidation(type="list", formula1='"Yes,No,Unclear"', allow_blank=True)
dv_noc = DataValidation(type="list", formula1='"Valid,Expired,Not Obtained,Unknown"', allow_blank=True)
dv_invest = DataValidation(type="list", formula1='"Ongoing,Completed,Unknown"', allow_blank=True)
dv_conf = DataValidation(type="list", formula1='"Confirmed - Official Report,Confirmed - Multiple News Sources,Single Source - Unverified,Conflicting Reports"', allow_blank=True)
dv_cause = DataValidation(type="list", formula1=f"=Causes_Taxonomy!$D$2:$D${CAUSES_LAST_ROW}", allow_blank=True)
dv_product = DataValidation(type="list", formula1=f"=Products_Taxonomy!$B$2:$B${PROD_LAST_ROW}", allow_blank=True)

MAXROW = 500
for dv in (dv_hospital_type, dv_severity, dv_oxy, dv_noc, dv_invest, dv_conf, dv_cause, dv_product):
    ws.add_data_validation(dv)

dv_hospital_type.add(f"{col['Hospital_Type']}2:{col['Hospital_Type']}{MAXROW}")
dv_severity.add(f"{col['Severity_Category']}2:{col['Severity_Category']}{MAXROW}")
dv_cause.add(f"{col['Root_Cause_Category']}2:{col['Root_Cause_Category']}{MAXROW}")
dv_product.add(f"{col['Product_Equipment_Involved']}2:{col['Product_Equipment_Involved']}{MAXROW}")
dv_oxy.add(f"{col['Oxygen_Related']}2:{col['Oxygen_Related']}{MAXROW}")
dv_noc.add(f"{col['Fire_Safety_NOC_Status']}2:{col['Fire_Safety_NOC_Status']}{MAXROW}")
dv_invest.add(f"{col['Investigation_Status']}2:{col['Investigation_Status']}{MAXROW}")
dv_conf.add(f"{col['Data_Confidence']}2:{col['Data_Confidence']}{MAXROW}")

# ---------------------------------------------------------------------------
# Failure_Analysis
# ---------------------------------------------------------------------------
ws = wb.create_sheet("Failure_Analysis")
fa_headers = ["Analysis_ID", "Title", "Product_Failure_Mode_Cluster", "Linked_Incident_IDs", "Number_of_Incidents",
              "Narrative", "Standards_Referenced", "Recommended_Mitigation", "Training_Implications",
              "Date_Created", "Last_Updated"]
fa_rows = [
    ("FA-001", "Aged/overloaded electrical distribution as the dominant hospital fire trigger",
     "Electrical Distribution - Short Circuit",
     "HF-2026-001, HF-2026-002, HF-2026-003, HF-2026-004, HF-2026-005, HF-2026-006", 6,
     "All six hospital fires logged so far in 2026 - two fatal (Cuttack/Odisha govt medical college ICU; Muzaffarpur/Bihar private hospital ICU) and four near-misses or minor-injury events with no deaths (Kanpur/UP govt hospital electric panel; Rewa/MP National Hospital OPD corridor; Bengaluru/Karnataka private hospital basement duct; Rewa/MP Gandhi Memorial Hospital SNCU ward) - were attributed by initial reports to electrical short circuits. The near-miss cases matter as much as the fatal ones here: they show the same failure mechanism (short circuit in wiring/panel/duct) recurring across government and private facilities, ICUs, general wards, OPD corridors, basements and a neonatal unit, with outcome apparently determined less by the ignition cause itself than by how fast staff detected it and how quickly patients could be moved - i.e. detection and evacuation readiness, not the electrical fault, is the variable separating a contained incident from a fatal one. This mirrors the broader, well-documented pattern that electrical short circuits from aged or poorly maintained wiring are the leading cause of building fires in India generally, and hospitals are especially exposed because: (1) ICUs, SNCUs and diagnostic areas carry dense, often retrofitted electrical loads (monitors, ventilators, imaging, AC, incubators) on wiring not originally sized for that load; (2) 24/7 continuous operation prevents routine de-energized inspection/maintenance windows; (3) panels and ducts are frequently sited in or adjacent to occupied clinical space rather than isolated fire-rated rooms; (4) two incidents (HF-2026-004, HF-2026-006) occurred in the same city (Rewa, MP) nine days apart, a possible signal of a localized infrastructure-age or maintenance-practice issue worth a targeted follow-up once more sources are available. Few of the six incidents' source reports specify the exact failed component (cable, breaker, joint, or an overloaded circuit) - this is the key open question for future investigation reports to close.",
     "IS 732 (electrical wiring), IS 8623 (switchgear assemblies), National Building Code 2016 Part 4 (Fire and Life Safety) - panel siting, circuit segregation and fire compartmentation of electrical rooms.",
     "1) Mandate third-party electrical safety audits (thermal imaging of panels/joints/ducts) at fixed intervals, prioritizing ICUs, SNCUs, basements and other continuous-load areas. 2) Require dedicated, correctly-rated circuits for high-draw ICU/diagnostic/neonatal equipment rather than shared/retrofitted circuits. 3) Site distribution panels and cable ducts in fire-rated, ventilated enclosures separated from patient-occupied space where structurally feasible; where not feasible, require local smoke/heat detection and automatic isolation. 4) Track panel/wiring age against rated service life and force replacement rather than run-to-failure. 5) Treat near-miss reports (no injury/fatality) as leading indicators, not non-events - log and review every one for the same rigor as a fatal incident.",
     "Facility engineering/maintenance staff: panel and duct inspection checklists, thermal-imaging interpretation, load-balancing basics. Nursing/ICU/SNCU staff: recognizing early signs of electrical fault (burning smell, flickering, warm switch plates, tripping breakers) and immediate escalation protocol. All staff: the Kanpur (HF-2026-003) and Bengaluru (HF-2026-005) cases are usable positive-outcome training examples - both contained rapidly by staff action (extinguishers; fast evacuation before fire services arrived) with zero casualties despite 100+ and 21 people on-site respectively, illustrating that a well-drilled response converts a serious electrical fault into a near miss rather than a fatality.",
     "2026-09-13", "2026-09-13"),
    ("FA-002", "Neonatal (NICU/SNCU) units as a high-consequence fire cluster, regardless of root cause",
     "Multiple (Electrical, HVAC, Medical Equipment) - Neonatal Unit Location",
     "HF-2026-006, HF-2026-009, HF-2026-012, HF-2026-014", 4,
     "Four incidents in this database originated specifically in neonatal care areas (SNCU/NICU), and their root causes differ - an electrical short circuit (HF-2026-006, Rewa SNCU), an AC-unit malfunction (HF-2026-009, Chhatrapati Sambhajinagar NICU), an internal fault in a neonatal thermal unit/ventilator (HF-2026-012, Amravati NICU) and an AC short circuit and explosion (HF-2026-014, Adilabad RIMS SNCU) - yet all four sit in the same distinctive risk category: rooms where every occupant is a newborn who cannot self-evacuate, walk, or even be quickly identified/carried without dedicated equipment (bassinet, incubator, warmer), and where bed-for-bed powered-device density (warmers, ventilators, monitors, phototherapy units) is far higher than a general ward. Of the four, two were fatal - Amravati (HF-2026-012), killing 3 of 9 infants present, and Adilabad (HF-2026-014), where 5 premature infants died of smoke inhalation, the deadliest incident logged since Cuttack; the other two evacuated every infant without fire injuries (two extremely premature twins evacuated in HF-2026-009 later died at the receiving hospital, which clinicians attributed to prematurity rather than smoke, but it shows that even a well-executed NICU evacuation is a risk for the most fragile infants). Adilabad adds an egress-design finding: the unit's doors reportedly locked after the blast and rescuers had to break the glass, and the infants died of smoke rather than flame, so smoke control and doors that open without power or a key matter as much as suppression. The distinguishing factor was not the root cause (all four are plausible, common equipment/electrical faults) but the interval between fault onset and staff intervention, and whether the unit's specific evacuation drill (moving neonates as a group, with their life-support equipment) had been rehearsed. This reframes the mitigation question: general electrical-safety audits (FA-001) reduce the frequency of a NICU fire starting, but a NICU-specific rapid-evacuation capability is what determines whether one becomes fatal.",
     "IEC 60601-1 (medical electrical equipment - general safety), IEC 60601-2-19/-2-20/-2-21 (particular requirements for infant incubators, infant transport incubators, and infant radiant warmers), National Building Code 2016 Part 4 (fire and life safety, area-specific evacuation planning).",
     "1) Preventive-maintenance and biomedical-engineering inspection schedule specifically for NICU/SNCU thermal units, warmers, ventilators and dedicated AC units, with fault-code and unusual-heat logging separate from general facility maintenance. 2) NICU-specific horizontal-evacuation drills at fixed intervals - practicing moving multiple neonates with attached equipment as a unit, not just an ambulatory-patient walk-through. 3) Continuous smoke/heat detection at the bassinet-cluster level inside NICUs/SNCUs, not only room-level detection. 4) A standing mutual-transfer arrangement with a nearby facility for emergency neonatal overflow, as used ad hoc in both HF-2026-009 and HF-2026-012. 5) Segregate high-power NICU equipment onto per-bay circuits so one device or AC-unit fault trips locally rather than affecting the whole unit. 6) Make every NICU/SNCU exit door open from inside without power, key or breaking glass, and add smoke extraction or a smoke-free refuge route, since smoke inhalation killed the Adilabad infants.",
     "NICU/SNCU nursing staff: simultaneous multi-infant evacuation drills (carry protocols, mobile-bassinet drills), recognizing device fault indicators (unusual heat, smell, fault lights) on warmers/ventilators/incubators, and immediate device power-isolation without waiting for facility electricians. Facility engineers/biomedical staff: NICU-specific preventive-maintenance checklists distinct from general building electrical audits. Hospital administration: pre-arranged emergency-transfer agreements with nearby facilities for neonatal overflow.",
     "2026-09-13", "2026-09-28"),
    ("FA-003", "Room air-conditioning units as a recurring ignition source in critical patient-care areas",
     "Air Conditioning/HVAC - AC Unit Short Circuit/Overheating",
     "HF-2026-009, HF-2026-013, HF-2026-014 (confirmed AC fault); HF-2026-011 (AC suspected, unresolved)", 4,
     "Three incidents in this database are attributed to a fault in a room air-conditioning unit, and a fourth lists it as one of two unresolved hypotheses: the fatal SNCU fire at RIMS Adilabad (HF-2026-014; a window AC reportedly burst into flames after a short circuit and its plastic parts fed dense smoke that killed 5 premature infants), the NICU fire at Asian Hospital, Chhatrapati Sambhajinagar (HF-2026-009, hospital director confirmed an AC malfunction; 58 patients evacuated), the anaesthesia-room fire beside an operation theatre at MKCG Medical College, Berhampur (HF-2026-013; AC destroyed, theatre empty), and the fourth-floor fire at Navodaya Hospital, Bareilly (HF-2026-011; AC short circuit vs. discarded smoking material, not established). All three confirmed cases sit in the highest-consequence spaces in a hospital - two neonatal units and an operating-theatre complex - where AC units run continuously and are typically fitted in ceilings or walls beside patient beds or anaesthesia equipment. ACs combine a compressor and capacitor under sustained load, dust-clogged coils and, in older sites, wiring that was not sized for them, which is a plausible route to overheating and arcing; no report identifies the failed component or brand, so this cluster cannot yet say whether the cause is maintenance, installation quality or product defect. Adilabad is the first case with any asset or maintenance detail, and it is not yet verified: Telangana Today (S031) cites unnamed officials and an unnamed expert for 130+ ACs bought in 2016, no regular in-house electricians, missed quarterly inspections, delayed repairs, alleged sub-standard refrigerant and compressor overload, and a six-member committee is examining the AC and electrical system, so its report may be the first to yield product or maintenance evidence. In HF-2026-013 the response was undermined by an expired extinguisher and staff untrained in using it, and the hospital's fire-safety certificates had reportedly lapsed for several blocks since 2023; the MKCG fire history (four fires in 2026 per S021) suggests a site where electrical and AC faults recur because the underlying maintenance and certification gaps were never closed.",
     "IS 1391 (room air conditioners), IS 732 (electrical wiring), IEC 60335-2-40 (safety of air conditioners), National Building Code 2016 Part 4 (fire and life safety); fire-extinguisher servicing per IS 2190.",
     "1) Put every AC serving a NICU, OT/anaesthesia room, ICU or dialysis unit on a documented preventive-maintenance schedule (capacitor and compressor checks, coil cleaning, electrical connection inspection) with the service history kept on file. 2) Give each critical-area AC its own correctly rated breaker with overload/thermal cut-out, so a fault trips locally rather than the whole ward. 3) Record brand, model and installation date for each AC when a fire occurs, so future investigations can test whether one product line is over-represented. 4) Add smoke detection near AC units in NICU/OT spaces. 5) Audit extinguisher validity tags monthly and retrain staff, since the extinguisher is the first response when an AC fire starts. 6) Track the age of the installed AC fleet (Adilabad's units were reportedly bought in 2016) and retire units on a service-life schedule instead of running them until they fail. 7) Put a named, dedicated electrician or biomedical technician in charge of critical-area AC and electrical upkeep rather than relying on outsourced staff assigned other duties.",
     "Facility engineers: AC preventive-maintenance checklists and dedicated-circuit requirements for critical areas. Nursing, OT and anaesthesia staff: smoke or burning smell from an AC is a fire precursor, so isolate the unit and raise the alarm immediately rather than waiting to see whether it clears. All clinical staff: hands-on extinguisher use, with drills that confirm the extinguisher at the point of use is in date.",
     "2026-09-18", "2026-09-28"),
]
last = write_table(ws, fa_headers, fa_rows, widths=[11, 34, 26, 26, 10, 90, 45, 60, 60, 13, 13])
FA_LAST_ROW = last

# ---------------------------------------------------------------------------
# Training_Recommendations
# ---------------------------------------------------------------------------
ws = wb.create_sheet("Training_Recommendations")
tr_headers = ["Rec_ID", "Linked_Cause_Category", "Recommendation", "Target_Audience", "Priority",
              "Linked_Incidents_Analyses"]
tr_rows = [
    ("TR-001", "Electrical - Short Circuit", "Train facility engineering staff to run and log periodic thermal-imaging inspections of all distribution panels, ducts and high-load circuits, with mandatory escalation on any hotspot found.", "Facility Engineers/Maintenance", "High", "FA-001; HF-2026-001 to HF-2026-006 (all 6 incidents)"),
    ("TR-002", "Electrical - Short Circuit", "Train ICU/ward/SNCU nursing staff to recognize early electrical-fault warning signs (burning smell, warm sockets, flickering lights, tripping breakers) and to immediately alert engineering and initiate the local evacuation-readiness check.", "Nursing/Clinical Staff", "High", "FA-001; HF-2026-001, HF-2026-002, HF-2026-006"),
    ("TR-003", "Structural/Fire Safety", "Run quarterly fire-extinguisher and evacuation drills specifically simulating bed-bound/ICU/neonatal patient transfer, not just ambulatory evacuation.", "Nursing/Clinical Staff, Security", "High", "HF-2026-001, HF-2026-002, HF-2026-006"),
    ("TR-004", "Oxygen System", "Train staff handling portable O2 cylinders and MGPS on safe storage distance from heat/electrical sources, valve inspection, and emergency shut-off procedure, given oxygen-enriched atmospheres sharply worsen any nearby ignition event.", "Nursing/Clinical Staff, Biomedical/Facility Engineers", "High", "FA-001 (contextual - no confirmed oxygen-cause incident logged yet)"),
    ("TR-005", "Structural/Fire Safety", "Use the Kanpur (HF-2026-003) and Bengaluru (HF-2026-005) incidents as case studies for rapid first-response fire suppression and evacuation: staff contained/evacuated ahead of fire services, with zero casualties among 100+ and 21 people present respectively.", "All Staff", "Medium", "HF-2026-003, HF-2026-005"),
    ("TR-006", "Structural/Fire Safety", "Establish and incentivize a near-miss reporting culture: every contained fire/electrical-fault event, however minor, should be logged and reviewed with the same rigor as a fatal incident, since 10 of the 15 incidents logged to date caused no injury but shared the same or a closely related root cause as the fatal ones. MKCG Berhampur (HF-2026-013) shows the cost of not doing so: four fires in 2026 at one hospital, with lapsed fire-safety certificates and expired extinguishers still in place after the earlier ones.", "All Staff, Hospital Administration", "High", "FA-001; HF-2026-003, HF-2026-004, HF-2026-005, HF-2026-006, HF-2026-007, HF-2026-008, HF-2026-009, HF-2026-010, HF-2026-013, HF-2026-015"),
    ("TR-007", "Medical Equipment", "Run NICU/SNCU-specific fire-response drills that rehearse moving multiple neonates with attached life-support equipment as a group, distinct from general ambulatory-patient evacuation drills; pair with a biomedical preventive-maintenance schedule for warmers/ventilators/incubators and NICU-dedicated AC units.", "NICU/SNCU Nursing Staff, Biomedical/Facility Engineers, Hospital Administration", "High", "FA-002; HF-2026-006, HF-2026-009, HF-2026-012, HF-2026-014"),
    ("TR-008", "HVAC/Cooling - AC Unit Overheating/Short Circuit", "Put every AC serving a NICU, operating theatre/anaesthesia room, ICU or dialysis unit on a documented preventive-maintenance schedule with its own rated breaker and thermal cut-out, and train nursing/OT staff that smoke or a burning smell from an AC is a fire precursor requiring immediate isolation and an alarm, not a wait-and-see.", "Facility Engineers/Maintenance, Nursing/OT/Anaesthesia Staff", "High", "FA-003; HF-2026-009, HF-2026-013, HF-2026-014"),
    ("TR-009", "Structural/Fire Safety", "Check extinguisher validity tags monthly and track fire-safety certificate expiry dates for every block; give all clinical staff hands-on extinguisher training, since HF-2026-013 found an expired extinguisher at the point of use and staff who reportedly did not know how to operate one.", "All Staff, Hospital Administration, Facility Engineers", "High", "FA-003; HF-2026-013"),
    ("TR-010", "Structural/Fire Safety - Blocked Fire Exits/Poor Evacuation Design", "Check that every NICU/SNCU and maternity-ward exit door opens from the inside without power, key or breaking glass, and drill staff and security on releasing it in the dark and in smoke. At Adilabad the doors reportedly locked after the blast, rescuers broke the glass, and the infants died of smoke inhalation.", "Facility Engineers, NICU/SNCU Nursing Staff, Security", "High", "FA-002; HF-2026-014"),
    ("TR-011", "HVAC/Cooling - AC Unit Overheating/Short Circuit", "Assign a named in-house electrician or biomedical technician to critical-area AC and electrical upkeep, with quarterly documented inspections and an equipment-age register, instead of relying on outsourced staff given other tasks. Reports on Adilabad allege exactly this gap (unverified until the inquiry reports), so treat it as a check to run in every hospital rather than an established finding.", "Hospital Administration, Facility Engineers/Maintenance", "High", "FA-003; HF-2026-014"),
]
last = write_table(ws, tr_headers, tr_rows, widths=[10, 22, 70, 30, 10, 40])

# ---------------------------------------------------------------------------
# Dashboard
# ---------------------------------------------------------------------------
ws = wb.create_sheet("Dashboard")
ws["A1"] = "Hospital Fire Database - Summary Dashboard"
ws["A1"].font = TITLE_FONT
ws["A2"] = "All figures below are formulas over the Incidents sheet and update automatically as rows are added."
ws["A2"].font = SUBTITLE_FONT
ws.merge_cells("A2:D2")

def kpi(row, label, formula):
    ws.cell(row=row, column=1, value=label).font = Font(name=FONT, bold=True, size=10)
    c = ws.cell(row=row, column=2, value=formula)
    c.font = Font(name=FONT, size=12, bold=True, color="1F4E5F")

SC = {h: f"Incidents!${l}$2:${l}$500" for h, l in col.items()}

kpi(4, "Total incidents logged", f"=COUNTA({SC['Incident_ID']})")
kpi(5, "Total fatalities", f"=SUM({SC['Fatalities']})")
kpi(6, "Total injuries", f"=SUM({SC['Injuries']})")
kpi(7, "Near-miss incidents (no injury/fatality)", f'=COUNTIF({SC["Severity_Category"]},"Near Miss (No Injury)")')
kpi(8, "Incidents with oxygen involvement (Yes)", f'=COUNTIF({SC["Oxygen_Related"]},"Yes")')
kpi(9, "Incidents still under investigation", f'=COUNTIF({SC["Investigation_Status"]},"Ongoing")')

ws["A11"] = "Incidents by Severity"
ws["A11"].font = Font(name=FONT, bold=True, size=11)
ws["A12"] = "Severity_Category"
ws["B12"] = "Incident Count"
style_header(ws, 12, 2)
severities = ["Fatal", "Non-Fatal Injury", "Near Miss (No Injury)", "Property/Equipment Damage Only", "Unknown"]
r = 13
for s in severities:
    ws.cell(row=r, column=1, value=s).font = BODY_FONT
    ws.cell(row=r, column=2, value=f'=COUNTIF({SC["Severity_Category"]},A{r})').font = BODY_FONT
    r += 1

ws["A20"] = "Incidents by State"
ws["A20"].font = Font(name=FONT, bold=True, size=11)
ws["A21"] = "State"
ws["B21"] = "Incident Count"
ws["C21"] = "Fatalities"
style_header(ws, 21, 3)
states = ["Odisha", "Bihar", "Uttar Pradesh", "Madhya Pradesh", "Karnataka", "Maharashtra", "Chhattisgarh", "Telangana", "Delhi", "West Bengal", "Tamil Nadu", "Other"]
r = 22
for s in states:
    ws.cell(row=r, column=1, value=s).font = BODY_FONT
    ws.cell(row=r, column=2, value=f'=COUNTIF({SC["State"]},A{r})').font = BODY_FONT
    ws.cell(row=r, column=3, value=f'=SUMIF({SC["State"]},A{r},{SC["Fatalities"]})').font = BODY_FONT
    r += 1

r += 1
ws.cell(row=r, column=1, value="Incidents by Root Cause Category").font = Font(name=FONT, bold=True, size=11)
r += 1
ws.cell(row=r, column=1, value="Root_Cause_Category")
ws.cell(row=r, column=2, value="Incident Count")
style_header(ws, r, 2)
r += 1
for i in range(2, CAUSES_LAST_ROW + 1):
    ws.cell(row=r, column=1, value=f"=Causes_Taxonomy!D{i}").font = BODY_FONT
    ws.cell(row=r, column=2, value=f'=COUNTIF({SC["Root_Cause_Category"]},A{r})').font = BODY_FONT
    r += 1

set_widths(ws, [34, 18, 14, 14])

wb.calculation.fullCalcOnLoad = True

wb.save("Hospital_Fire_Database_India.xlsx")
print("saved")
