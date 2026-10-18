import fs from "node:fs/promises";
import AdmZip from "adm-zip";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = "/Users/nathaliefutolan/Documents/ChatGPT/Decarby/outputs/decarby_awards_2027";
const outputPath = `${outputDir}/Decarby_Awards_Pipeline_2027.xlsx`;
const font = "Arial";
const navy = "#17324D";
const teal = "#0F6B68";
const sky = "#DDEBF7";
const paleTeal = "#DDEFEA";
const paleGold = "#FFF2CC";
const paleRed = "#FCE4D6";
const paleGrey = "#F2F4F7";
const green = "#E2F0D9";
const white = "#FFFFFF";
const text = "#1F2937";
const border = "#CFD6DE";

const d = (y, m, day) => new Date(Date.UTC(y, m - 1, day));

const awards = [
  {
    rec: "Target 2027", award: "Bluebell Sustainable Business Awards", category: "Energy Transition; Sustainable Small Business; Innovation Excellence", candidate: "Decarby", scope: "ACT",
    close2026: d(2026, 8, 10), window2027: "Expected Jul-Aug 2027; use 10 Aug as planning assumption", certainty: "Forecast from 2026 cycle",
    status: "Likely eligible", eligibility: "ACT business or organisation. Entry should show measurable sustainability or climate outcomes; select one strongest category.",
    requirements: "Choose one category; provide a project narrative, measurable energy/emissions outcomes, supporting calculations, approved photos and customer/partner evidence.",
    fit: 5, elig: 5, ready: 3, win: 4, marketing: 5, effort: 2, fee: "No fee stated",
    rationale: "Best local fit. Decarby's integrated solar, battery, electrification and energy-efficiency work aligns directly. Main gap is quantified customer and operational impact.",
    next: "Choose one flagship ACT project and establish before/after energy, cost and emissions metrics by 31 Mar 2027.",
    source: "https://www.climatechoices.act.gov.au/events-news/news/nominations-are-now-open-for-the-2026-bluebell-sustainable-business-awards"
  },
  {
    rec: "Target 2027", award: "Canberra Women in Business Awards", category: "Sustainability Business Woman; Business Development; Leadership", candidate: "Nana Sashiwa", scope: "ACT / Canberra region",
    close2026: d(2026, 9, 16), window2027: "Expected Jul-Sep 2027; use mid-Sep as planning assumption", certainty: "Forecast from 2026 cycle",
    status: "Needs confirmation", eligibility: "Award entry is a CWB member benefit. Candidate should be a woman leading or contributing in business in the Canberra region; confirm category rules when released.",
    requirements: "Hold or obtain CWB membership; complete category questions; show Nana's attributable leadership, client wins and business results; prepare references and current profile assets.",
    fit: 5, elig: 4, ready: 3, win: 4, marketing: 5, effort: 2, fee: "CWB membership: $150/year (individual)",
    rationale: "Strong local profile fit for Nana's business development and operations role. Sustainability and leadership categories provide credible paths without requiring founder status, subject to 2027 rules.",
    next: "Confirm Nana's CWB membership and capture revenue growth, client wins, process improvements and leadership examples quarterly.",
    source: "https://www.cwb.org.au/awards/"
  },
  {
    rec: "Target 2027", award: "Australian Trades Small Business Champion Awards", category: "Solar", candidate: "Decarby", scope: "Australia",
    close2026: d(2026, 1, 31), window2027: "Likely Jan-Feb 2027; monitor from Dec 2026", certainty: "Forecast; 2026 site contains inconsistent event dates",
    status: "Likely eligible", eligibility: "Australian trade retail/service business with an ABN and 40 or fewer FTE; business must fit the Solar category.",
    requirements: "Confirm ABN and FTE count; complete business/category questions; supply team photo and supporting documents; evidence install quality, growth and customer satisfaction.",
    fit: 5, elig: 5, ready: 3, win: 3, marketing: 4, effort: 2, fee: "Entry fee not stated; verify on launch",
    rationale: "Direct Solar category and small-business field make this more targeted than broad national programs. Customer proof, install quality and growth evidence will decide competitiveness.",
    next: "Prepare a reusable business profile, team photo, three customer case studies and quality/compliance evidence before entries open.",
    source: "https://www.championawards.com.au/trades/categories"
  },
  {
    rec: "Target 2027", award: "Women in Industry Awards", category: "Business Development Success of the Year; Excellence in Energy; Women in Leadership", candidate: "Nana Sashiwa", scope: "Australia",
    close2026: d(2026, 4, 2), window2027: "Expected Feb-Apr 2027", certainty: "Forecast from 2026 cycle",
    status: "Likely eligible", eligibility: "Nominee must reside and work in Australia. Business Development criteria focus on initiatives, clients won and financial benefit delivered in the two years before close.",
    requirements: "Address the published criteria with initiatives, clients won and financial benefit from the prior two years; add independent evidence, portrait/bio and finalist availability.",
    fit: 5, elig: 5, ready: 2, win: 2, marketing: 5, effort: 3, fee: "Not stated; verify",
    rationale: "Excellent role/category match, but a national field and explicit financial evidence make it demanding. Nana needs attributable two-year growth results.",
    next: "Create a Nana impact log now: initiative, her role, client or partnership won, revenue or pipeline value, and proof source.",
    source: "https://womeninindustry.com.au/nominations/"
  },
  {
    rec: "Target 2027", award: "AIRAH National Awards", category: "Best Retrofit Project; Industry Advancement & Inclusion", candidate: "Bilal Babadagi / Decarby", scope: "Australia",
    close2026: d(2026, 8, 31), window2027: "Expected Jun-Aug 2027", certainty: "Forecast from 2026 cycle",
    status: "Likely eligible", eligibility: "HVAC&R and building-services projects or individuals. Best Retrofit needs a completed upgrade with clear performance and sustainability outcomes; category rules may change.",
    requirements: "Submit a completed retrofit/initiative with design rationale, installation and commissioning details, measured performance and sustainability outcomes, technical evidence and photos.",
    fit: 5, elig: 4, ready: 2, win: 3, marketing: 5, effort: 4, fee: "Not stated; verify",
    rationale: "Bilal's AIRAH affiliation and building-services background are strong. A measured heat-pump, AC or whole-home electrification retrofit would be more competitive than a standard solar install.",
    next: "Nominate one technically distinctive retrofit and arrange baseline, commissioning and post-install performance measurement.",
    source: "https://awards.airah.org.au/2026-award-categories/"
  },
  {
    rec: "Stretch 2027", award: "Engineers Australia Excellence Awards", category: "Professional Engineer of the Year", candidate: "Bilal Babadagi", scope: "Australia; local-to-national pathway",
    close2026: d(2026, 7, 1), window2027: "Expected 1 Apr-1 Jul 2027", certainty: "Forecast from 2025-2026 pattern",
    status: "Likely eligible", eligibility: "Engineers Australia membership or engineering-community membership. Four equally weighted criteria cover transformational impact, judgement/innovation, ethics/sustainability and influence on the profession.",
    requirements: "Confirm current membership; prepare four criterion responses, a 2-3 page CV, short bio, portrait, pronunciation audio and referee details; verify current word limits.",
    fit: 5, elig: 5, ready: 3, win: 1, marketing: 5, effort: 4, fee: "No fee stated",
    rationale: "Bilal's CPEng and NER credentials support eligibility, but this is a highly competitive individual award requiring impact beyond ordinary business delivery and influence on the profession.",
    next: "Map Bilal's top three engineering achievements to the four criteria; quantify system-level impact and external influence.",
    source: "https://www.engineersaustralia.org.au/publications/excellence-awards-2026-people-awards-entry-handbook"
  },
  {
    rec: "Stretch 2027", award: "Smart Energy Excellence Awards", category: "Smart Installer; Emerging Leader; Smart Employer", candidate: "Decarby / eligible team member", scope: "Australia",
    close2026: null, window2027: "Likely Q1 2027; subscribe and verify", certainty: "Date not published in source reviewed",
    status: "Needs confirmation", eligibility: "Renewable-energy-sector program. 2026 categories included Smart Installer, Emerging Leader and Smart Employer; confirm membership and 2027 category rules.",
    requirements: "Confirm membership and entry conditions; quantify install quality, innovation, safety, employer practices and customer outcomes; prepare evidence and high-resolution images.",
    fit: 5, elig: 3, ready: 2, win: 2, marketing: 5, effort: 3, fee: "Membership/entry conditions TBC",
    rationale: "Strong sector relevance and direct installer category, but 2026 attracted more than 250 submissions and Decarby needs scale, measured quality and a differentiated installer story.",
    next: "Join the organiser mailing list; document install volume, defect/rework rate, customer ratings, safety and innovative designs.",
    source: "https://smartenergy.org.au/smart-energy-excellence-awards-2026-celebrates-industry-leaders-at-gala-dinner-in-sydney/"
  },
  {
    rec: "Stretch 2027", award: "Energy Efficiency Council Awards", category: "Residential energy-management project; Innovation; Champion", candidate: "Decarby / Bilal", scope: "Australia",
    close2026: null, window2027: "Likely Q1 2027; verify when nominations launch", certainty: "Cycle confirmed; 2027 date not published",
    status: "Likely eligible", eligibility: "Projects must demonstrate excellence in energy efficiency, electrification or demand flexibility. Individual champion categories generally require third-party nomination.",
    requirements: "Provide measured project outcomes, innovation and energy/electrification evidence; arrange third-party nomination for champion categories and representative availability if shortlisted.",
    fit: 4, elig: 4, ready: 2, win: 2, marketing: 4, effort: 4, fee: "TBC",
    rationale: "A strong fit only if Decarby can show measured whole-home or commercial electrification outcomes beyond a standard install.",
    next: "Design measurement and verification into one 2026-27 project, including energy baseline, bills, peak demand and comfort outcomes.",
    source: "https://eec.org.au/events/eec-award-winners-2026-draft/"
  },
  {
    rec: "Stretch 2027", award: "Banksia National Sustainability Awards", category: "Small to Medium Business Leadership; Business Services and Solutions; Built Environment", candidate: "Decarby", scope: "Australia",
    close2026: null, window2027: "Monitor from Apr 2027; closing date TBC", certainty: "2026 close was still TBA on official page",
    status: "Likely eligible", eligibility: "National sustainability program aligned to the UN SDGs. Entries need measurable environmental and social outcomes, a structured submission, case study and PR assets.",
    requirements: "Map results to relevant SDGs; complete the structured submission; prepare a 500-word case study, short summaries, three approved high-resolution photos and entry fee.",
    fit: 4, elig: 5, ready: 2, win: 1, marketing: 5, effort: 5, fee: "$150+GST if under 5 staff; $450+GST if under 20 staff/turnover under $10m (2026 NSW fee guide; verify national fee)",
    rationale: "Prestigious but evidence-heavy and national. Better after Decarby has a full year of verified impact and stronger whole-of-business sustainability metrics.",
    next: "Align impact measures to relevant SDGs and build a 500-word flagship case study with approved high-resolution photos.",
    source: "https://banksiafdn.com/about-awards/"
  },
  {
    rec: "Stretch 2027", award: "Telstra Best of Business Awards", category: "Change", candidate: "Decarby", scope: "Australia",
    close2026: d(2026, 8, 7), window2027: "Expected Jul-Aug 2027", certainty: "Forecast from 2026 cycle; categories changed in 2026",
    status: "Likely eligible", eligibility: "Australian small or medium business; 2026 public guidance indicated fewer than 200 staff and one category per entry. Confirm full 2027 rules.",
    requirements: "Choose one category; confirm workforce limit; document growth, purpose, customer impact and business outcomes; complete the staged application and verification materials.",
    fit: 4, elig: 4, ready: 2, win: 1, marketing: 5, effort: 5, fee: "No fee identified; verify",
    rationale: "High marketing value but extremely broad competition. Decarby needs strong growth, purpose and customer-impact proof before investing heavily.",
    next: "Use 2027 as a readiness year unless revenue growth, customer count and quantified emissions impact become exceptional by Q2.",
    source: "https://www.telstra.com.au/small-business/best-of-business-awards"
  },
  {
    rec: "Monitor 2027", award: "Australian Small Business Champion Awards", category: "Electrical Business; Specialised Small Business; Business Growth", candidate: "Decarby", scope: "Australia",
    close2026: d(2026, 3, 2), window2027: "Likely opens Nov 2026 and closes around Mar 2027", certainty: "Forecast from 2026 cycle",
    status: "Likely eligible", eligibility: "Australian service business with 40 or fewer FTE and a current ABN. One industry category plus permitted optional categories; New Business is limited to under two years.",
    requirements: "Confirm ABN/FTE eligibility; select one industry category; complete category questions; provide a team photo, business results and supporting documents.",
    fit: 4, elig: 5, ready: 3, win: 2, marketing: 4, effort: 3, fee: "Free in 2026",
    rationale: "Good small-business recognition, though less solar-specific than the Trades program. Use only if the application can be largely reused.",
    next: "Monitor from Nov 2026 and reuse the Trades award evidence pack if the categories and timing remain complementary.",
    source: "https://www.championawards.com.au/business/terms-and-conditions"
  },
  {
    rec: "Monitor 2027", award: "ACT Women's Awards", category: "Excellence in Industry", candidate: "Nana Sashiwa", scope: "ACT region",
    close2026: d(2026, 5, 10), window2027: "Expected Mar-May 2027", certainty: "Forecast from 2026 cycle",
    status: "Likely eligible", eligibility: "Women, gender-diverse and non-binary people in the ACT region whose achievements create lasting positive change. Self- or third-party nomination was allowed in 2026.",
    requirements: "Complete the nomination form; give specific examples of lasting positive change and ACT impact; include a referee and supporting evidence where permitted.",
    fit: 4, elig: 5, ready: 2, win: 2, marketing: 4, effort: 3, fee: "No fee stated",
    rationale: "Good local platform, but criteria emphasise lasting positive change beyond routine job performance. Stronger after Nana has a visible industry or community initiative.",
    next: "Identify a Nana-led initiative with external impact, such as consumer education, inclusion, or a partnership that improves access to clean energy.",
    source: "https://www.act.gov.au/community/honours-and-awards/act-womens-awards-and-honour-roll"
  },
  {
    rec: "Monitor 2027", award: "Clean Energy Council Awards", category: "Community Value and Impact; Equity, Inclusion and Diversity; Communications", candidate: "Decarby / partner consortium", scope: "Australia",
    close2026: null, window2027: "Likely Q2 2027; verify", certainty: "2027 dates not published",
    status: "Not ready", eligibility: "Category-specific national awards for demonstrated initiatives, partnerships or campaigns in the clean-energy sector. Ordinary commercial delivery alone is unlikely to be enough.",
    requirements: "Enter a substantial initiative or partnership; quantify community, inclusion or communications impact; secure partner endorsements and a publishable case study.",
    fit: 3, elig: 3, ready: 1, win: 1, marketing: 5, effort: 5, fee: "TBC",
    rationale: "Prestigious sector award but current categories favour major community, inclusion, collaboration or communications initiatives rather than small installer performance.",
    next: "Pursue only with a substantial partner-led community or inclusion program and measurable multi-stakeholder outcomes.",
    source: "https://cleanenergycouncil.org.au/events/awards"
  },
  {
    rec: "Conditional", award: "Australian Women's Small Business Champion Awards", category: "Trade Services; Professional Services; Influential Woman", candidate: "Nana / Decarby", scope: "Australia",
    close2026: d(2026, 7, 31), window2027: "Expected Jun-Jul 2027", certainty: "Forecast from 2026 cycle",
    status: "Needs confirmation", eligibility: "Business must have a woman owner or at least 50% woman ownership, operate in Australia, hold an ABN and meet employee limits. Entry was free in 2026.",
    requirements: "Verify at least 50% woman ownership, ABN and employee limits; complete category questions; provide team photo and supporting PDF evidence.",
    fit: 4, elig: 1, ready: 2, win: 2, marketing: 4, effort: 3, fee: "Free in 2026",
    rationale: "Relevant only if Decarby has at least 50% woman ownership. Nana's current public role is manager, so ownership must be confirmed before any work begins.",
    next: "Confirm company ownership. If the 50% requirement is not met, remove this award from the active pipeline.",
    source: "https://www.championawards.com.au/womens/terms-and-conditions"
  },
  {
    rec: "Conditional", award: "AIRAH Towards Net Zero Competition", category: "David Norman Sustainability Award", candidate: "Ahmad Bilal or another eligible team member", scope: "Australia",
    close2026: d(2026, 9, 7), window2027: "If repeated: registration around Sep; staged submissions follow", certainty: "Future recurrence not confirmed",
    status: "Needs confirmation", eligibility: "Entrant was required to be under 35 on 1 Jul 2026 and work or study in sustainability, engineering, building services or related fields. Technical written and modelling stages; $150 fee in 2026.",
    requirements: "Confirm age and relevant work/study; pay registration fee; complete the entrant's own written knowledge assessment and, if advanced, building-energy simulation stage.",
    fit: 5, elig: 2, ready: 2, win: 2, marketing: 4, effort: 5, fee: "$150 in 2026",
    rationale: "Strong technical alignment but age, modelling capability and recurrence need confirmation. Rules prohibit AI-written submissions, so the entrant must do the competition work personally.",
    next: "Confirm candidate age and building-energy simulation experience; monitor AIRAH for a 2027 edition.",
    source: "https://awards.airah.org.au/towards-net-zero-competition/"
  },
  {
    rec: "Conditional", award: "ACT Multicultural Awards", category: "Outstanding Excellence Award for Diversity and Inclusion", candidate: "Decarby / team initiative", scope: "ACT / Canberra region",
    close2026: null, window2027: "Monitor ACT awards calendar in 2027", certainty: "2027 date not published",
    status: "Not ready", eligibility: "ACT or Canberra-region nominee. The 2026 category focused on people in the community sector supporting cultural diversity and inclusion.",
    requirements: "Show an ACT/Canberra community-sector initiative with sustained multicultural inclusion outcomes, beneficiary evidence and credible partner or community support.",
    fit: 2, elig: 2, ready: 1, win: 1, marketing: 3, effort: 3, fee: "No fee stated",
    rationale: "Team diversity alone is insufficient. A sustained community-facing program with measurable multicultural inclusion outcomes would be needed.",
    next: "Do not apply unless Decarby creates or partners on a genuine community inclusion initiative with evidence of benefit.",
    source: "https://www.act.gov.au/community/honours-and-awards/multicultural-awards"
  },
  {
    rec: "Conditional", award: "AusMumpreneur Awards", category: "Service Business; Sustainability-related category; Business Excellence", candidate: "Nana Sashiwa", scope: "Australia / New Zealand",
    close2026: d(2026, 7, 10), window2027: "Expected nomination by early Jul; application later in Jul", certainty: "Forecast from 2026 cycle",
    status: "Needs confirmation", eligibility: "Entrant must be an Australian or New Zealand mother who owns a business, nonprofit or blog. Paid categories started at $150 in 2026.",
    requirements: "Privately confirm parent and ownership status; select and pay for category; complete application questions and prepare required links, banners and media assets.",
    fit: 3, elig: 1, ready: 2, win: 2, marketing: 3, effort: 3, fee: "From $150 in 2026",
    rationale: "Only relevant if Nana is both a mother and an owner of Decarby or another eligible business. Do not assume either fact.",
    next: "Confirm personal and ownership eligibility privately before retaining this in the pipeline.",
    source: "https://ausmumpreneur.com/entry/"
  },
  {
    rec: "Conditional", award: "Australian Ladies in Business Initiative Awards", category: "Earth Keeper; Professional Service; Social Change Maker", candidate: "Nana Sashiwa", scope: "Australia",
    close2026: d(2026, 8, 3), window2027: "Expected Jun-Aug 2027", certainty: "Forecast from 2026 cycle",
    status: "Needs confirmation", eligibility: "Open to Australian women in business; 2026 rules required an Australian business with ABN and membership of the organiser's free online Facebook group.",
    requirements: "Confirm Nana personally meets business/ABN rules; join the required online group; pay per category; submit separate category responses and evidence.",
    fit: 3, elig: 2, ready: 2, win: 3, marketing: 2, effort: 2, fee: "$99 per category in 2026",
    rationale: "Lower strategic prestige than the local CWB and national Women in Industry programs. Consider only if Nana personally meets the business/ABN rules and capacity remains.",
    next: "Keep as a backup; prioritise CWB and Women in Industry first.",
    source: "https://alibiawards.com.au/entry-requirements/"
  },
  {
    rec: "Conditional", award: "Engineers Australia Excellence Awards", category: "Professional or Emerging Professional Engineer of the Year", candidate: "Ahmad Bilal", scope: "Australia; local-to-national pathway",
    close2026: d(2026, 7, 1), window2027: "Expected 1 Apr-1 Jul 2027", certainty: "Forecast from 2025-2026 pattern",
    status: "Needs confirmation", eligibility: "Engineers Australia or engineering-community membership is required. Emerging category is limited to age 35 or under at the awards date.",
    requirements: "Confirm membership and age/category; prepare four criterion responses, a 2-3 page CV, short bio, portrait, pronunciation audio and referee details.",
    fit: 4, elig: 2, ready: 1, win: 1, marketing: 4, effort: 4, fee: "No fee stated",
    rationale: "Potentially relevant, but Ahmad's membership, credentials, age, role and evidence are not available in the project brief.",
    next: "Confirm Ahmad's full name, role, EA membership, credentials, age and strongest quantified engineering achievements.",
    source: "https://www.engineersaustralia.org.au/publications/excellence-awards-2026-people-awards-entry-handbook"
  }
];

const evidenceRows = [
  ["Eligibility", "Company profile", "ABN, incorporation date, ownership percentages, FTE count, locations and service areas", "Nana", d(2026, 11, 13), "Not started", "Unlocks small-business, women-owned and new-business eligibility checks"],
  ["Eligibility", "Individual profiles", "Age, role, employment dates, memberships, licences, CPEng/NER/SAA/AIRAH details for each candidate", "Each candidate", d(2026, 11, 13), "Not started", "Needed before assigning individual categories"],
  ["Business results", "Growth scorecard", "Monthly revenue growth, sales pipeline, new clients, conversion rate, repeat/referral share and staff growth", "Nana", d(2026, 11, 30), "Not started", "Core evidence for Nana, Telstra and small-business awards"],
  ["Customer impact", "Impact register", "For each completed project: kW solar, kWh storage, annual kWh saved/generated, bill savings, peak reduction and estimated tCO2-e", "Bilal / Ahmad", d(2026, 12, 15), "Not started", "Use consistent calculation method and keep supporting bills or model files"],
  ["Case studies", "Three flagship projects", "One residential whole-home electrification, one commercial project and one technically complex retrofit", "Bilal", d(2027, 1, 31), "Not started", "Include problem, Decarby's decisions, complexity, quantified outcomes and photos"],
  ["Customer proof", "Testimonials and permissions", "Named testimonials, review links, client reference contacts and written permission to publish project details/photos", "Nana", d(2027, 1, 31), "Not started", "Avoid confidential claims that cannot be published"],
  ["Quality", "Delivery performance", "Install volume, on-time completion, defects, rework, callbacks, safety incidents and warranty response time", "Operations", d(2027, 2, 15), "Not started", "Strong differentiator for installer and business awards"],
  ["Engineering", "Bilal achievement map", "Three achievements mapped to EA impact, judgement, ethical stewardship and professional influence criteria", "Bilal", d(2027, 2, 28), "Not started", "Use measurable outcomes and third-party proof"],
  ["Diversity", "Inclusion evidence", "Workforce profile, hiring pathways, mentoring, inclusive policies, supplier diversity and outcomes", "Nana", d(2027, 3, 15), "Not started", "Do not claim diversity impact without a program and measured results"],
  ["Sustainability", "Decarby operations baseline", "Office electricity, fleet, travel, waste, procurement and improvement targets", "Nana", d(2027, 3, 31), "Not started", "Needed for whole-of-business sustainability categories"],
  ["Community", "Education and access initiative", "Document workshops, consumer education, partnerships, pro bono work or equitable clean-energy access", "Team", d(2027, 4, 30), "Not started", "Supports ACT Women, CEC, Banksia and diversity categories"],
  ["External proof", "Referees and endorsements", "At least three clients/partners plus one industry referee; capture titles, contact details and approved statements", "Nana", d(2027, 5, 15), "Not started", "Prioritise independent and senior referees"],
  ["Media assets", "Awards asset pack", "100-word bios, company boilerplate, team and portrait photos, project images, logos and name-pronunciation audio", "Marketing", d(2027, 5, 31), "Not started", "Reuse across applications; keep current consent records"],
  ["Governance", "Claims register", "Every quantitative claim linked to a source file, owner, reporting period and approval status", "Nana", d(2027, 6, 15), "Not started", "Prevents inconsistent claims across applications"]
];

const timelineRows = [
  ["Oct-Nov 2026", "Set foundations", "Confirm eligibility facts; subscribe to all organiser updates; assign one evidence owner; create impact and growth registers."],
  ["Dec 2026-Jan 2027", "First application pack", "Prepare Australian Trades / Small Business Champion material; select three flagship projects; obtain client permissions."],
  ["Feb-Mar 2027", "Nana and energy awards", "Draft Women in Industry entry; assess Smart Energy and EEC launches; complete business metrics and operations baseline."],
  ["Apr-May 2027", "Individual and local impact", "Decide ACT Women's entry; draft Bilal's Engineers Australia criteria; build community/inclusion evidence."],
  ["Jun-Jul 2027", "National submissions", "Submit Engineers Australia if evidence threshold is met; prepare Telstra, Banksia and women's small-business conditional entries."],
  ["Jul-Sep 2027", "Local and industry peak", "Submit Bluebell, AIRAH and Canberra Women in Business; reuse verified evidence and tailor category narratives."],
  ["Oct-Dec 2027", "Review and leverage", "Record results and judge feedback; publish finalist/winner marketing; update evidence gaps and 2028 pipeline."]
];

const wb = Workbook.create();
const summary = wb.worksheets.add("Priority plan");
const register = wb.worksheets.add("Award register");
const evidence = wb.worksheets.add("Evidence plan");

for (const s of [summary, register, evidence]) {
  s.showGridLines = false;
  s.getRange("A1:W200").format.font = { name: font, size: 10, color: text };
}
summary.tabColor = navy;
register.tabColor = teal;
evidence.tabColor = "#5B9BD5";

// Priority plan
summary.getRange("A2:H2").merge();
summary.getRange("A2").values = [["Decarby awards pipeline 2027"]];
summary.getRange("A2:H2").format.font = { name: font, size: 16, bold: true, color: navy };
summary.getRange("A3:H3").merge();
summary.getRange("A3").values = [["Prepared 19 October 2026. Closing windows for 2027 are forecasts until organisers publish their next cycle."]];
summary.getRange("A3:H3").format.font = { name: font, size: 10, italic: true, color: "#5F6B7A" };
summary.getRange("A4:H4").format.borders = { bottom: { style: "thin", color: navy } };

summary.getRange("A6:B6").values = [["Recommended 2027 targets", null]];
summary.getRange("D6:E6").values = [["Awards screened", null]];
summary.getRange("G6:H6").values = [["Current 2026 status", null]];
summary.getRange("A6:H6").format.font = { name: font, size: 10, bold: true, color: navy };
summary.getRange("A7:B8").merge();
summary.getRange("A7").formulas = [["=COUNTIFS('Award register'!$B$5:$B$23,\"Target 2027\")"]];
summary.getRange("D7:E8").merge();
summary.getRange("D7").formulas = [["=COUNTA('Award register'!$D$5:$D$23)"]];
summary.getRange("G7:H8").merge();
summary.getRange("G7").values = [["All researched 2026 application windows have closed"]];
for (const rg of ["A7:B8", "D7:E8", "G7:H8"]) {
  summary.getRange(rg).format.fill = paleTeal;
  summary.getRange(rg).format.font = { name: font, size: rg === "G7:H8" ? 10 : 16, bold: true, color: navy };
  summary.getRange(rg).format.horizontalAlignment = "center";
  summary.getRange(rg).format.verticalAlignment = "center";
  summary.getRange(rg).format.wrapText = true;
  summary.getRange(rg).format.borders = { preset: "outside", style: "thin", color: border };
}

summary.getRange("A10:H10").values = [["Rank", "Award", "Candidate", "Expected close", "Priority score", "Est. win probability", "Decision", "Immediate focus"]];
summary.getRange("A10:H10").format = { fill: navy, font: { name: font, size: 10, bold: true, color: white }, horizontalAlignment: "center", verticalAlignment: "center", wrapText: true, borders: { insideVertical: { style: "thin", color: white } } };
for (let i = 0; i < 5; i++) {
  const sr = 11 + i;
  const rr = 5 + i;
  summary.getRange(`A${sr}:H${sr}`).formulas = [[
    `=${i + 1}`, `='Award register'!D${rr}`, `='Award register'!F${rr}`, `='Award register'!I${rr}`,
    `='Award register'!S${rr}`, `='Award register'!T${rr}`, `='Award register'!B${rr}`, `='Award register'!W${rr}`
  ]];
}
summary.getRange("A11:H15").format.borders = { bottom: { style: "thin", color: border } };
summary.getRange("A11:A15").format.horizontalAlignment = "center";
summary.getRange("E11:F15").format.horizontalAlignment = "center";
summary.getRange("E11:E15").format.numberFormat = "0";
summary.getRange("F11:F15").format.numberFormat = "0%";
summary.getRange("B11:D15").format.wrapText = true;
summary.getRange("G11:H15").format.wrapText = true;
summary.getRange("A11:H15").format.verticalAlignment = "top";

summary.getRange("A16:H16").merge();
summary.getRange("A16").values = [["Quick use: open Award register and filter Recommendation, Eligibility status, Candidate, Expected 2027 close or Estimated win probability. Each Official source cell is a clickable link."]];
summary.getRange("A16:H16").format = { fill: sky, font: { name: font, size: 9, italic: true, color: navy }, wrapText: true, verticalAlignment: "center", borders: { preset: "outside", style: "thin", color: border } };

summary.getRange("A18:C18").merge();
summary.getRange("A18").values = [["Recommended sequencing"]];
summary.getRange("A18:C18").format = { fill: teal, font: { name: font, size: 11, bold: true, color: white }, verticalAlignment: "center" };
summary.getRange("A19:C19").values = [["Period", "Focus", "Actions"]];
summary.getRange("A19:C19").format = { fill: sky, font: { name: font, size: 10, bold: true, color: navy }, horizontalAlignment: "center", borders: { bottom: { style: "thin", color: navy } } };
summary.getRange("A20").write(timelineRows);
summary.getRange(`A20:C${19 + timelineRows.length}`).format.wrapText = true;
summary.getRange(`A20:C${19 + timelineRows.length}`).format.verticalAlignment = "top";
summary.getRange(`A20:C${19 + timelineRows.length}`).format.borders = { bottom: { style: "thin", color: border } };

summary.getRange("E18:H18").merge();
summary.getRange("E18").values = [["Before starting any application"]];
summary.getRange("E18:H18").format = { fill: teal, font: { name: font, size: 11, bold: true, color: white }, verticalAlignment: "center" };
summary.getRange("E19:H26").values = [
  ["1", "Confirm Nana's ownership status", "Several entrepreneur/women-owned awards depend on this", "Owner: Nana / directors"],
  ["2", "Confirm Ahmad's identity and credentials", "EA membership, age, role and technical record are unknown", "Owner: Ahmad"],
  ["3", "Select three flagship projects", "Prioritise measurable electrification and retrofit outcomes", "Owner: Bilal"],
  ["4", "Start monthly growth tracking", "Revenue, clients, conversion, referrals and pipeline", "Owner: Nana"],
  ["5", "Measure customer impact consistently", "Energy, bills, peak demand and emissions", "Owner: Bilal / Ahmad"],
  ["6", "Secure publication permissions", "Customer names, photos, data and testimonials", "Owner: Nana"],
  ["7", "Choose one category per program", "Do not dilute the evidence across weak categories", "Owner: Project team"],
  ["8", "Apply only above the evidence threshold", "At least three quantified outcomes and two independent proof points", "Owner: Project team"]
];
summary.getRange("E19:H26").format.wrapText = true;
summary.getRange("E19:E26").format.horizontalAlignment = "center";
summary.getRange("E19:H26").format.verticalAlignment = "top";
summary.getRange("E19:H26").format.borders = { bottom: { style: "thin", color: border } };
summary.getRange("E19:E26").format.fill = paleGold;

summary.getRange("A29:H29").merge();
summary.getRange("A29").values = [["Scoring note"]];
summary.getRange("A29:H29").format.font = { name: font, size: 11, bold: true, color: navy };
summary.getRange("A30:H31").merge();
summary.getRange("A30").values = [["Priority score weights: strategic fit 25%, eligibility confidence 20%, evidence readiness 15%, winability 15%, marketing value 20%, and lower effort 5%. Indicative likelihood is a capped planning estimate based on fit, eligibility, readiness and competition; it is not an organiser-published probability. Scores should be updated as evidence and 2027 rules become available."]];
summary.getRange("A30:H31").format = { fill: paleGrey, font: { name: font, size: 9, italic: true, color: "#5F6B7A" }, wrapText: true, verticalAlignment: "top", borders: { preset: "outside", style: "thin", color: border } };
summary.getRange("A33:H33").merge();
summary.getRange("A33").values = [["Automated monitoring"]];
summary.getRange("A33:H33").format = { fill: teal, font: { name: font, size: 11, bold: true, color: white } };
summary.getRange("A34:H35").merge();
summary.getRange("A34").values = [["GitHub checks all official award sources every Friday and Monday at 6:00 AM PHT. Monday's run emails an HTML briefing and this workbook to nathalie@decarbysolar.com.au through the existing authorised Outlook connection."]];
summary.getRange("A34:H35").format = { fill: paleTeal, font: { name: font, size: 10, color: navy }, wrapText: true, verticalAlignment: "center", borders: { preset: "outside", style: "thin", color: border } };

// Award register
register.getRange("A2:Y2").merge();
register.getRange("A2").values = [["Award register and fit assessment"]];
register.getRange("A2:Y2").format.font = { name: font, size: 16, bold: true, color: navy };
register.getRange("A3:Y3").merge();
register.getRange("A3").values = [["Use the filter arrows in row 4. Scores are editable inputs (1=low, 5=high); probability is an internal planning estimate. 2027 dates remain forecasts until verified."]];
register.getRange("A3:Y3").format.font = { name: font, size: 10, italic: true, color: "#5F6B7A" };
const headers = ["Rank", "Recommendation", "Eligibility status", "Award / program", "Best category", "Candidate", "Scope", "Last confirmed close", "Expected 2027 close / window", "Date certainty", "Eligibility / criteria summary", "Requirements / evidence needed", "Fit", "Eligibility confidence", "Evidence readiness", "Winability", "Marketing value", "Effort", "Priority score", "Estimated win probability", "Fee / membership", "Why it fits / gap", "Next action", "Official source", "Source checked"];
register.getRange("A4:Y4").values = [headers];
register.getRange("A4:Y4").format = { fill: navy, font: { name: font, size: 10, bold: true, color: white }, horizontalAlignment: "center", verticalAlignment: "center", wrapText: true, borders: { insideVertical: { style: "thin", color: white } } };
const regRows = awards.map((a, idx) => [idx + 1, a.rec, a.status, a.award, a.category, a.candidate, a.scope, a.close2026, a.window2027, a.certainty, a.eligibility, a.requirements, a.fit, a.elig, a.ready, a.win, a.marketing, a.effort, null, null, a.fee, a.rationale, a.next, null, d(2026, 10, 19)]);
register.getRange("A5").write(regRows);
for (let i = 0; i < awards.length; i++) {
  const row = 5 + i;
  register.getRange(`S${row}`).formulas = [[`=ROUND((M${row}*25+N${row}*20+O${row}*15+P${row}*15+Q${row}*20+(6-R${row})*5)/5,0)`]];
  register.getRange(`T${row}`).formulas = [[`=ROUND((M${row}*0.30+N${row}*0.25+O${row}*0.20+P${row}*0.25)/5*40%,2)`]];
  const safeUrl = awards[i].source.replaceAll('"', '""');
  register.getRange(`X${row}`).formulas = [[`=HYPERLINK("${safeUrl}","Open official page")`]];
}
const regEnd = 4 + awards.length;
register.getRange(`A5:Y${regEnd}`).format.verticalAlignment = "top";
register.getRange(`B5:L${regEnd}`).format.wrapText = true;
register.getRange(`U5:Y${regEnd}`).format.wrapText = true;
register.getRange(`A5:A${regEnd}`).format.horizontalAlignment = "center";
register.getRange(`M5:T${regEnd}`).format.horizontalAlignment = "center";
register.getRange(`H5:H${regEnd}`).format.numberFormat = "dd mmm yyyy";
register.getRange(`Y5:Y${regEnd}`).format.numberFormat = "dd mmm yyyy";
register.getRange(`T5:T${regEnd}`).format.numberFormat = "0%";
register.getRange(`A5:Y${regEnd}`).format.borders = { bottom: { style: "thin", color: border } };
register.getRange(`M5:R${regEnd}`).format.fill = paleGold;
register.getRange(`S5:T${regEnd}`).format.fill = paleTeal;
register.getRange(`M5:R${regEnd}`).dataValidation = { rule: { type: "whole", operator: "between", formula1: 1, formula2: 5 } };
register.getRange(`B5:B${regEnd}`).dataValidation = { rule: { type: "list", values: ["Target 2027", "Stretch 2027", "Monitor 2027", "Conditional"] } };
register.getRange(`C5:C${regEnd}`).dataValidation = { rule: { type: "list", values: ["Likely eligible", "Needs confirmation", "Not ready"] } };
register.getRange(`B5:B${regEnd}`).conditionalFormats.add("containsText", { text: "Target 2027", format: { fill: green, font: { bold: true, color: "#375623" } } });
register.getRange(`B5:B${regEnd}`).conditionalFormats.add("containsText", { text: "Stretch 2027", format: { fill: paleGold, font: { color: "#7F6000" } } });
register.getRange(`B5:B${regEnd}`).conditionalFormats.add("containsText", { text: "Conditional", format: { fill: paleRed, font: { color: "#9C0006" } } });
register.getRange(`C5:C${regEnd}`).conditionalFormats.add("containsText", { text: "Likely eligible", format: { fill: green, font: { color: "#375623" } } });
register.getRange(`C5:C${regEnd}`).conditionalFormats.add("containsText", { text: "Needs confirmation", format: { fill: paleGold, font: { color: "#7F6000" } } });
register.getRange(`C5:C${regEnd}`).conditionalFormats.add("containsText", { text: "Not ready", format: { fill: paleRed, font: { color: "#9C0006" } } });
register.getRange(`S5:S${regEnd}`).conditionalFormats.add("colorScale", { colors: ["#F8696B", "#FFEB84", "#63BE7B"], thresholds: ["min", { type: "percentile", value: 50 }, "max"] });
register.getRange(`T5:T${regEnd}`).conditionalFormats.add("dataBar", { color: "#5B9BD5", gradient: true });
const regTable = register.tables.add(`A4:Y${regEnd}`, true, "AwardsRegister");
regTable.style = "TableStyleMedium2";
regTable.showBandedColumns = false;
register.freezePanes.freezeRows(4);
register.freezePanes.freezeColumns(4);

// Evidence plan
evidence.getRange("A2:G2").merge();
evidence.getRange("A2").values = [["Evidence plan for 2027 applications"]];
evidence.getRange("A2:G2").format.font = { name: font, size: 16, bold: true, color: navy };
evidence.getRange("A3:G3").merge();
evidence.getRange("A3").values = [["Build the proof once, verify it, then tailor the narrative to each award. Dates below are internal preparation deadlines."]];
evidence.getRange("A3:G3").format.font = { name: font, size: 10, italic: true, color: "#5F6B7A" };
evidence.getRange("A5:G5").values = [["Evidence area", "Deliverable", "What to capture", "Owner", "Internal due", "Status", "Why it matters"]];
evidence.getRange("A5:G5").format = { fill: navy, font: { name: font, size: 10, bold: true, color: white }, horizontalAlignment: "center", verticalAlignment: "center", wrapText: true, borders: { insideVertical: { style: "thin", color: white } } };
evidence.getRange("A6").write(evidenceRows);
const evEnd = 5 + evidenceRows.length;
evidence.getRange(`A6:G${evEnd}`).format.wrapText = true;
evidence.getRange(`A6:G${evEnd}`).format.verticalAlignment = "top";
evidence.getRange(`A6:G${evEnd}`).format.borders = { bottom: { style: "thin", color: border } };
evidence.getRange(`E6:E${evEnd}`).format.numberFormat = "dd mmm yyyy";
evidence.getRange(`F6:F${evEnd}`).dataValidation = { rule: { type: "list", values: ["Not started", "In progress", "Complete", "Blocked"] } };
evidence.getRange(`F6:F${evEnd}`).conditionalFormats.add("containsText", { text: "Complete", format: { fill: green, font: { bold: true, color: "#375623" } } });
evidence.getRange(`F6:F${evEnd}`).conditionalFormats.add("containsText", { text: "Blocked", format: { fill: paleRed, font: { bold: true, color: "#9C0006" } } });
evidence.getRange(`E6:F${evEnd}`).format.fill = paleGold;
const evTable = evidence.tables.add(`A5:G${evEnd}`, true, "EvidencePlan");
evTable.style = "TableStyleMedium2";
evidence.freezePanes.freezeRows(5);

evidence.getRange("A23:G23").merge();
evidence.getRange("A23").values = [["Submission quality gate"]];
evidence.getRange("A23:G23").format = { fill: teal, font: { name: font, size: 11, bold: true, color: white } };
evidence.getRange("A24:G28").values = [
  ["1", "Eligibility confirmed", "Candidate and company meet every published rule", null, null, null, null],
  ["2", "Three quantified outcomes", "At least three results with source files and reporting periods", null, null, null, null],
  ["3", "Two independent proof points", "Client, partner, certification, media or referee evidence", null, null, null, null],
  ["4", "Category-specific story", "Clear problem, Decarby's distinctive role, complexity, action and lasting result", null, null, null, null],
  ["5", "Publication approval", "All names, photos, figures and confidential material cleared before submission", null, null, null, null]
];
evidence.getRange("A24:C28").format.wrapText = true;
evidence.getRange("A24:A28").format.horizontalAlignment = "center";
evidence.getRange("A24:C28").format.borders = { bottom: { style: "thin", color: border } };

// Sizing
summary.getRange("A1:H40").format.rowHeight = 18;
summary.getRange("A2:H2").format.rowHeight = 28;
summary.getRange("A3:H3").format.rowHeight = 22;
summary.getRange("A10:H10").format.rowHeight = 34;
summary.getRange("A11:H15").format.rowHeight = 54;
summary.getRange("A20:C26").format.rowHeight = 82;
summary.getRange("E19:H26").format.rowHeight = 58;
summary.getRange("A30:H31").format.rowHeight = 38;
summary.getRange("A34:H35").format.rowHeight = 28;
summary.getRange("A:A").format.columnWidth = 10;
summary.getRange("B:B").format.columnWidth = 30;
summary.getRange("C:C").format.columnWidth = 22;
summary.getRange("D:D").format.columnWidth = 25;
summary.getRange("E:E").format.columnWidth = 12;
summary.getRange("F:F").format.columnWidth = 15;
summary.getRange("G:G").format.columnWidth = 16;
summary.getRange("H:H").format.columnWidth = 42;

register.getRange("A1:Y30").format.rowHeight = 18;
register.getRange("A2:Y2").format.rowHeight = 28;
register.getRange("A3:Y3").format.rowHeight = 22;
register.getRange("A4:Y4").format.rowHeight = 54;
register.getRange(`A5:Y${regEnd}`).format.rowHeight = 96;
const widths = [7,14,16,30,32,22,16,15,29,20,48,48,7,12,12,8,11,7,12,14,28,48,46,22,14];
for (let i = 0; i < widths.length; i++) register.getRangeByIndexes(0, i, 30, 1).format.columnWidth = widths[i];

evidence.getRange("A1:G35").format.rowHeight = 18;
evidence.getRange("A2:G2").format.rowHeight = 28;
evidence.getRange("A3:G3").format.rowHeight = 22;
evidence.getRange("A5:G5").format.rowHeight = 36;
evidence.getRange(`A6:G${evEnd}`).format.rowHeight = 55;
const evWidths = [16,24,58,18,15,14,42];
for (let i = 0; i < evWidths.length; i++) evidence.getRangeByIndexes(0, i, 35, 1).format.columnWidth = evWidths[i];

wb.recalculate();

const summaryCheck = await wb.inspect({ kind: "table", range: "Priority plan!A1:H35", include: "values,formulas", tableMaxRows: 35, tableMaxCols: 8 });
console.log(summaryCheck.ndjson);
const registerCheck = await wb.inspect({ kind: "table", range: `Award register!A1:Y${regEnd}`, include: "values,formulas", tableMaxRows: 10, tableMaxCols: 25 });
console.log(registerCheck.ndjson);
const errorCheck = await wb.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!", options: { useRegex: true, maxResults: 100 }, summary: "final formula error scan" });
console.log(errorCheck.ndjson);

await fs.mkdir(outputDir, { recursive: true });
for (const [sheetName, fileName, range] of [
  ["Priority plan", "preview_priority.png", "A1:H35"],
  ["Award register", "preview_register.png", `A1:Y${regEnd}`],
  ["Evidence plan", "preview_evidence.png", "A1:G28"]
]) {
  const preview = await wb.render({ sheetName, range, scale: 1, format: "png" });
  await fs.writeFile(`${outputDir}/${fileName}`, new Uint8Array(await preview.arrayBuffer()));
}
const xlsx = await SpreadsheetFile.exportXlsx(wb);
await xlsx.save(outputPath);
await patchHyperlinkCache(outputPath);
console.log(JSON.stringify({ outputPath, sheets: ["Priority plan", "Award register", "Evidence plan"], awards: awards.length }));

async function patchHyperlinkCache(path) {
  const zip = new AdmZip(path);
  const entry = zip.getEntry("xl/worksheets/sheet2.xml");
  if (!entry) throw new Error("Award register worksheet XML was not found");
  let xml = entry.getData().toString("utf8");
  for (let row = 5; row <= regEnd; row++) {
    const cellPattern = new RegExp(`(<x:c r="X${row}"[^>]*?) t="e"(><x:f>HYPERLINK\\([\\s\\S]*?<\\/x:f><x:v>)[\\s\\S]*?(<\\/x:v><\\/x:c>)`);
    xml = xml.replace(cellPattern, `$1 t="str"$2Open official page$3`);
  }
  zip.updateFile("xl/worksheets/sheet2.xml", Buffer.from(xml, "utf8"));
  zip.writeZip(path);
}
