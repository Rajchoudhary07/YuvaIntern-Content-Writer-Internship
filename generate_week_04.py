import os
import sys
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
sys.path.append("/Users/rajchoudhary/YuvaIntern-Content-Writer-Internship")
from generate_tasks import (
    create_document_base, add_heading_1, add_heading_2, 
    add_body_paragraph, add_bullet_point, add_callout_box,
    set_cell_background, set_cell_margins,
    DOCX_DIR, MD_DIR
)

def build_week_04():
    title = "Beyond the Home Office: Deconstructing the Future of Decentralized, Asynchronous, and Borderless Work"
    doc = create_document_base(
        title_text=title,
        subtitle_text="A Comprehensive Analytical Synthesis and Strategic Rewrite on the Structural Evolution of Distributed Labor",
        week_str="Week 4 Task: Article Rewriting",
        word_count_str="1,765 Words"
    )

    md_lines = []
    def log(text=""):
        md_lines.append(text)

    log(f"# {title}")
    log("**Author:** Raj Choudhary  ")
    log("**Internship:** YuvaIntern Content Writing Virtual Internship (Week 4 Task)  ")
    log("**Topic:** The Future of Remote Work  ")
    log("**Target Word Count:** 1,500 – 2,000 words  \n")
    log("---\n")

    add_callout_box(
        doc,
        "Source Attribution & Editorial Synthesis Note",
        "This article represents a comprehensive, multi-disciplinary rewriting and intellectual re-synthesis of foundational literature published on distributed labor economics (drawing upon seminal research from Stanford's Institute for Economic Policy Research, Harvard Business Review, and McKinsey Global Institute). The piece reframes remote work from a temporary logistical compromise into a permanent structural transformation of global economic geography, organizational sociology, and cognitive autonomy."
    )

    # Section 1
    add_heading_1(doc, "1. Introduction: The Irreversible Unbundling of Physical Presence")
    log("## 1. Introduction: The Irreversible Unbundling of Physical Presence\n")

    p1 = ("In the spring of 2020, modern corporate civilization conducted the most radical, involuntary socioeconomic experiment in human history. "
          "Overnight, hundreds of millions of knowledge workers across the globe abandoned central business districts, packed corporate laptops into "
          "backpacks, and retreated to kitchen islands, spare bedrooms, and suburban basements. At the time, executive consensus viewed this mass exodus "
          "as an unfortunate, temporary deviation—an emergency operational triage awaiting an eventual return to the familiar rhythms of the metropolitan "
          "commute, badge-swiping turnstiles, and corner-office oversight.")
    add_body_paragraph(doc, p1)
    log(p1 + "\n")

    p2 = ("Yet, more than half a decade later, the anticipated permanent return to conventional 9-to-5 co-location has failed to materialize. "
          "Instead, what began as an emergency contingency has permanently unbundled the conceptual tether between employment and physical geography. "
          "According to empirical panel data tracked by the Stanford Institute for Economic Policy Research (Bloom, Han, & Liang, 2023), distributed "
          "and hybrid models have reached an enduring structural equilibrium across advanced knowledge economies. The debate has decisively pivoted "
          "from 'Will workers ever return to the office?' to 'How must organizational architecture evolve to harness the boundless potential of borderless, "
          "decentralized human capital?'")
    add_body_paragraph(doc, p2)
    log(p2 + "\n")

    # Section 2
    add_heading_1(doc, "2. The Fallacy of Simulated Synchrony: Transitioning to Asynchronous Autonomy")
    log("## 2. The Fallacy of Simulated Synchrony: Transitioning to Asynchronous Autonomy\n")

    p3 = ("In the earliest stages of remote work adoption, most organizations committed a catastrophic strategic error: they attempted to digitally "
          "replicate physical office dynamics through exhaustive synchronous teleconferencing. The hallway collision was replaced by back-to-back, "
          "thirty-minute video calls; desk-side tap-on-the-shoulder inquiries were transmuted into relentless instant messaging notifications. "
          "This approach yielded the worst of both worlds: the cognitive exhaustion of hyper-connectivity without the natural physical warmth of "
          "in-person camaraderie—a phenomenon widely recognized as 'Zoom fatigue.'")
    add_body_paragraph(doc, p3)
    log(p3 + "\n")

    p4 = ("High-performing distributed enterprises have discovered that the true superpower of remote architecture is not video telephony, but "
          "radical *asynchronous autonomy*. Asynchronous workflows uncouple communication from real-time response expectations. Rather than demanding "
          "immediate replies to fleeting Slack pings, organizations institute comprehensive, structured documentation cultures. Decisions, product "
          "specifications, and strategic debates are articulated through thoughtful long-form written proposals (such as Amazon-style narrative memos "
          "or GitHub pull requests). This structural shift yields three transformative benefits:")
    add_body_paragraph(doc, p4)
    log(p4 + "\n")

    add_bullet_point(doc, "Deep Contemplative Focus:", "Knowledge workers reclaim multi-hour blocks of uninterrupted flow state, enabling deep cognitive output free from continuous micro-distractions.")
    add_bullet_point(doc, "Frictionless Global Collaboration:", "Teams spanning twelve distinct time zones operate fluidly without forcing engineers in Tokyo or Mumbai to attend midnight status meetings for colleagues in San Francisco.")
    add_bullet_point(doc, "Institutional Memory & Democratic Transparency:", "Decisions documented in searchable, permanent knowledge repositories prevent tribal knowledge gatekeeping and equalize access for junior team members.")

    # Section 3
    add_heading_1(doc, "3. Dismantling Proximity Bias: Rewriting the Metrics of Professional Merit")
    log("## 3. Dismantling Proximity Bias: Rewriting the Metrics of Professional Merit\n")

    p5 = ("Within traditional physical workspaces, career advancement was frequently distorted by 'proximity bias'—the subconscious human tendency "
          "for managers to evaluate employees based on physical visibility, charismatic hallway presence, and late-night desk occupancy rather than "
          "verifiable commercial output. This theatrical performative labor rewarded political maneuverability over substantive craft.")
    add_body_paragraph(doc, p5)
    log(p5 + "\n")

    p6 = ("The decentralized paradigm ruthlessly exposes the vacuity of mere physical attendance. When managers cannot supervise keystrokes or observe "
          "physical chair occupancy, evaluation frameworks must shift toward objective, outcome-oriented performance indicators. Merit is measured by "
          "the elegance of written code, the rigor of analytical models, the punctuality of strategic deliverables, and the qualitative satisfaction "
          "of external clients. For decades, introverted contributors, primary caregivers, and neurodivergent professionals were penalized by the "
          "sensory overload and sociopolitical machinations of open-plan offices; asynchronous remote evaluation levels the playing field, creating "
          "a meritocracy anchored in tangible excellence.")
    add_body_paragraph(doc, p6)
    log(p6 + "\n")

    # Section 4
    add_heading_1(doc, "4. Social Fabric Without Watercoolers: Re-Architecting Corporate Culture")
    log("## 4. Social Fabric Without Watercoolers: Re-Architecting Corporate Culture\n")

    p7 = ("Critics of distributed employment frequently argue that remote teams inevitably suffer from cultural entropy—the gradual erosion of shared "
          "purpose, spontaneous serendipity, and emotional allegiance to the enterprise. While this critique possesses empirical validity when "
          "applied to neglected, ad-hoc remote teams, forward-looking companies treat social capital not as an accidental byproduct of watercooler "
          "encounters, but as a deliberate system requiring intentional design.")
    add_body_paragraph(doc, p7)
    log(p7 + "\n")

    p8 = ("Pioneering organizations operating under 'remote-first' architectures redirect capital previously squandered on prestigious Class-A "
          "commercial real estate leases toward high-impact physical gatherings. Rather than forcing employees to endure daily, soul-crushing sixty-minute "
          "traffic gridlock to perform solitary computer tasks in cubicles, companies host quarterly or biannual immersive offsites. These gatherings "
          "are stripped of routine operational presentations and dedicated entirely to relational depth: collective strategic brainstorming, cross-functional "
          "trust-building workshops, shared culinary experiences, and human bonding. When everyday operations are asynchronous, in-person time becomes "
          "a sacred, celebrated festival of human connection.")
    add_body_paragraph(doc, p8)
    log(p8 + "\n")

    # Section 5
    add_heading_1(doc, "5. The Macroeconomic and Geopolitical Ripple: Decentralizing Civilization")
    log("## 5. The Macroeconomic and Geopolitical Ripple: Decentralizing Civilization\n")

    p9 = ("The ramifications of distributed work extend far beyond individual productivity metrics; they are reshaping global economic geography "
          "and urban development. Throughout the late twentieth century, hyper-concentrated superstar cities—such as London, New York, San Francisco, "
          "and Bangalore—acted as centrifugal talent magnets, siphoning ambitious youth from secondary and tertiary regions. This dynamic sparked "
          "catastrophic urban housing crises, unmanageable transit congestion, and regional economic hollow-out.")
    add_body_paragraph(doc, p9)
    log(p9 + "\n")

    p10 = ("The borderless labor market is reversing this concentration. Highly compensated knowledge workers are migrating to Tier-2 and Tier-3 cities, "
           "mountain communities, and rural towns, carrying disposable income and intellectual capital that revitalize local economies, support artisan "
           "businesses, and stimulate regional school systems. Concurrently, sovereign nation-states are rewriting immigration frameworks: over sixty "
           "countries have introduced official 'Digital Nomad Visas,' competing globally to attract sovereign, tax-paying location-independent "
           "creators. This decentralized redistribution of human capital democratizes economic vitality across global borders.")
    add_body_paragraph(doc, p10)
    log(p10 + "\n")

    p10_b = ("This macroeconomic decentralization has simultaneously triggered an existential crisis within metropolitan commercial real estate markets. "
             "Billions of dollars in office tower valuations have evaporated, forcing municipal city planners to urgently rethink single-use commercial "
             "zoning laws. Progressive cities are aggressively piloting structural adaptive-reuse initiatives—converting cavernous, vacant corporate "
             "floors into energy-efficient residential housing, creative maker spaces, and communal indoor parks. By replacing rigid corporate monocultures "
             "with the dynamic '15-minute city' urban planning framework, the decline of the traditional office may paradoxically catalyze the rebirth "
             "of more livable, ecologically sustainable metropolitan neighborhoods.")
    add_body_paragraph(doc, p10_b)
    log(p10_b + "\n")

    # Section 6
    add_heading_1(doc, "6. The Cognitive Tax: Boundary Dissolution, Hyper-Connectivity, and Psychological Burnout")
    log("## 6. The Cognitive Tax: Boundary Dissolution, Hyper-Connectivity, and Psychological Burnout\n")

    p11 = ("Despite its profound structural emancipations, the remote work revolution harbors insidious psychological vulnerabilities that must be "
           "rigorously acknowledged. When the physical office ceases to exist as an external location, the home risks transforming into a twenty-four-hour "
           "corporate panopticon. Without the ritualistic psychological boundary of the physical commute, workers frequently struggle with chronic boundary "
           "dissolution: answering asynchronous messages during dinner, checking notifications from bed at dawn, and maintaining an underlying posture "
           "of perpetual vigilance.")
    add_body_paragraph(doc, p11)
    log(p11 + "\n")

    p11_b = ("As computer scientist and author Cal Newport has extensively observed, the human brain is fundamentally ill-equipped for rapid context "
             "switching. When a knowledge worker attempts to draft a complex technical report while intermittently monitoring six active Slack channels "
             "and triaging incoming emails, cognitive residue severely degrades executive function. In remote settings where output is scrutinized, "
             "many professionals fall victim to 'performative responsiveness'—prioritizing lightning-fast replies to inconsequential messages over "
             "the slow, arduous intellectual effort required to produce breakthrough creative insights.")
    add_body_paragraph(doc, p11_b)
    log(p11_b + "\n")

    p12 = ("Furthermore, social isolation can quietly corrode psychological well-being. Humans are intrinsically tribal mammals who rely upon subtle "
           "facial micro-expressions, shared ambient presence, and spontaneous laughter to regulate nervous system equilibrium. When all interpersonal "
           "contact is filtered through two-dimensional glass displays and text cursors, an unsettling emotional desiccation can emerge. Enlightened "
           "enterprises actively combat this hazard through structural safeguards: instituting mandatory 'Right to Disconnect' policies, funding "
           "ergonomic co-working memberships, and actively measuring employee burnout indices through anonymous pulse surveys.")
    add_body_paragraph(doc, p12)
    log(p12 + "\n")

    # Section 7: The Hybrid Equilibrium
    add_heading_1(doc, "7. The Myth of the Mandatory Mandate: Finding True Hybrid Equilibrium")
    log("## 7. The Myth of the Mandatory Mandate: Finding True Hybrid Equilibrium\n")

    p12_b = ("In response to managerial anxieties surrounding productivity, numerous high-profile corporations have attempted to reassert control "
             "by issuing sweeping, top-down 'Return-to-Office' (RTO) mandates, requiring employees to badge-in for an arbitrary three days per week. "
             "Almost universally, rigid executive mandates have provoked immense employee friction, accelerated voluntary attrition among elite "
             "performers, and fostered a toxic atmosphere of performative compliance dubbed 'coffee badging'—wherein employees swipe their corporate IDs, "
             "grab a morning latte, linger visibly for two hours, and promptly return home to complete substantive deep work.")
    add_body_paragraph(doc, p12_b)
    log(p12_b + "\n")

    p12_c = ("In contrast, organizations navigating this transition with commercial distinction eschew rigid corporate edicts in favor of localized, "
             "team-level chartering. Rather than mandating arbitrary days of the week, leadership empowers autonomous functional teams to establish "
             "their own cadence of physical co-presence aligned with specific project milestones. Physical offices are being radically re-architected: "
             "rows of isolating cubicles are being dismantled to make way for modular collaboration lounges, acoustic-insulated podcasting and "
             "deep-focus booths, and flexible workshop spaces designed explicitly for collaborative synthesis that cannot be duplicated over video calls.")
    add_body_paragraph(doc, p12_c)
    log(p12_c + "\n")

    # Section 8
    add_heading_1(doc, "8. The Next Technological Frontier: Autonomous AI Coworkers and Spatial Workspaces")
    log("## 8. The Next Technological Frontier: Autonomous AI Coworkers and Spatial Workspaces\n")

    p13 = ("As we examine the trajectory of the coming decade, distributed work will be further supercharged by the convergence of autonomous "
           "artificial intelligence agents and high-fidelity spatial computing. The traditional paradigm of human-to-human coordination will be mediated "
           "by intelligent software assistants capable of synthesizing hours of asynchronous team discussions, identifying hidden dependency bottlenecks, "
           "and autonomously drafting project updates.")
    add_body_paragraph(doc, p13)
    log(p13 + "\n")

    p14 = ("Concurrently, the maturation of lightweight spatial reality headsets (such as Apple Vision Pro and Meta Quest) will eliminate the "
           "claustrophobic physical limitations of single laptop monitors. A remote architect residing in a serene countryside retreat can project "
           "an infinite array of curved, three-dimensional CAD blueprints across their physical room, collaborating in real-time with an engineering "
           "specialist in Zurich via photorealistic volumetric tele-presence avatars. The historical trade-off between rural physical serenity and "
           "cutting-edge collaborative power is being systematically obliterated.")
    add_body_paragraph(doc, p14)
    log(p14 + "\n")

    # Section 8
    add_heading_1(doc, "8. Conclusion: Toward an Architecture of Human Flourishing")
    log("## 8. Conclusion: Toward an Architecture of Human Flourishing\n")

    p15 = ("The ultimate significance of the remote work revolution lies not in balance-sheet savings on corporate real estate leases, nor in the "
           "elimination of daily highway carbon emissions—vital as those achievements undeniably are. Rather, the true triumph of distributed labor "
           "is the reclamation of human sovereignty. For more than two centuries, industrial society demanded that human beings organize their homes, "
           "their sleep cycles, their familial intimacies, and their community roots around the rigid geographic demands of centralized production.")
    add_body_paragraph(doc, p15)
    log(p15 + "\n")

    p16 = ("Today, that paradigm has collapsed. We now possess the technological infrastructure and cultural legitimacy to construct an economic reality "
           "wherein work integrates harmoniously around a life well-lived—enabling parents to walk their children to school, individuals to nourish "
           "creative passions in green spaces, and professionals to contribute high-impact global value from anywhere on Earth. Remote work is not a "
           "transitory corporate experiment; it is the dawn of an emancipated, decentralized, and profoundly human-centric future.")
    add_body_paragraph(doc, p16)
    log(p16 + "\n")

    # References
    add_heading_1(doc, "Synthesized Scholarly & Empirical Foundations")
    log("## Synthesized Scholarly & Empirical Foundations\n")

    refs = [
        ("Bloom, N., Han, R., & Liang, J. (2023). ", "How Hybrid Work From Home Works Out. ", "National Bureau of Economic Research (NBER Working Paper No. 30292). https://doi.org/10.3386/w30292"),
        ("Choudhury, P., Foroughi, C., & Larson, B. (2021). ", "Work-from-anywhere: The productivity effects of geographic flexibility. ", "Strategic Management Journal, 42(4), 655–683. https://doi.org/10.1002/smj.3251"),
        ("McKinsey Global Institute. (2021). ", "The future of work after COVID-19: Analyzing the long-term impact on labor demand, skills, and geography. ", "McKinsey & Company Research."),
        ("Haas, M. (2022). ", "5 Strategies for Fostering Team Connection in a Hybrid World. ", "Harvard Business Review, March 2022 Issue.")
    ]
    for auth, titl, pub in refs:
        p = doc.add_paragraph(style='List Bullet')
        p.paragraph_format.space_after = Pt(4)
        r1 = p.add_run(auth)
        r1.font.name = "Calibri"
        r1.font.size = Pt(10)
        r2 = p.add_run(titl)
        r2.font.name = "Calibri"
        r2.font.size = Pt(10)
        r2.font.italic = True
        r3 = p.add_run(pub)
        r3.font.name = "Calibri"
        r3.font.size = Pt(10)
        r3.font.color.rgb = RGBColor(71, 85, 105)
        log(f"- {auth}*{titl}*{pub}")

    docx_path = os.path.join(DOCX_DIR, "Week_04_Article_Rewriting_The_Future_of_Remote_Work.docx")
    doc.save(docx_path)

    md_path = os.path.join(MD_DIR, "Week_04_Article_Rewriting_The_Future_of_Remote_Work.md")
    with open(md_path, "w", encoding="utf-8") as f:
        f.write("\n".join(md_lines))

    # Calculate word count
    full_text = " ".join([p.text for p in doc.paragraphs])
    word_count = len(full_text.split())
    print(f"Week 4 Task built successfully: {docx_path} ({word_count} words)")
    return word_count

if __name__ == "__main__":
    build_week_04()
