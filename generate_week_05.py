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

def build_week_05():
    title = "ZenithFit: Comprehensive Omnichannel Content Strategy & Growth Blueprint"
    doc = create_document_base(
        title_text=title,
        subtitle_text="A Data-Driven Content Architecture, Audience Persona Matrix, and Multi-Pillar Editorial Strategy for an AI Fitness App",
        week_str="Week 5 Task: Content Strategy Development",
        word_count_str="2,150 Words"
    )

    md_lines = []
    def log(text=""):
        md_lines.append(text)

    log(f"# {title}")
    log("**Brand:** ZenithFit (AI-Powered Adaptive Fitness & Wellness Companion)  ")
    log("**Author:** Raj Choudhary  ")
    log("**Internship:** YuvaIntern Content Writing Virtual Internship (Week 5 Task)  ")
    log("**Deliverable:** Content Strategy Architecture with 4 Pillars, 20+ Detailed Topics, Personas, Distribution Matrix, and KPI Scorecard  \n")
    log("---\n")

    add_callout_box(
        doc,
        "Executive Strategy Overview",
        "ZenithFit is an intelligent fitness application that bridges the gap between biomechanical science, behavioral psychology, and adaptive computer vision. This content strategy shifts user acquisition from expensive paid performance ads to authority-driven organic content. By democratizing elite athletic training principles into accessible, science-backed lifestyle habits, ZenithFit establishes category dominance across the global digital fitness market."
    )

    # Section 1
    add_heading_1(doc, "1. Executive Summary & Brand Positioning")
    log("## 1. Executive Summary & Brand Positioning\n")

    p1 = ("In a saturated digital wellness landscape dominated by superficial aesthetic influencers and generic workout timers, ZenithFit "
          "carves a distinct white space: 'Precision Longevity Fitness.' Powered by on-device computer-vision form tracking and real-time biometric "
          "feedback, ZenithFit delivers personalized, adaptive workout and recovery programs tailored to an individual's daily sleep deficit, "
          "heart rate variability (HRV), and muscle fatigue indices.")
    add_body_paragraph(doc, p1)
    log(p1 + "\n")

    p2 = ("The primary strategic objective of this content strategy is to position ZenithFit not merely as an app, but as an indispensable "
          "intellectual mentor. By generating high-utility educational content that solves immediate physical pain points—such as chronic lower "
          "back stiffness, training plateaus, and nutritional confusion—we build high-trust organic acquisition funnels that convert passive audience "
          "members into passionate, paying daily active subscribers.")
    add_body_paragraph(doc, p2)
    log(p2 + "\n")

    # Section 2: Target Audience Personas
    add_heading_1(doc, "2. Target Audience Personas (Customer Profiling)")
    log("## 2. Target Audience Personas (Customer Profiling)\n")

    personas = [
        ("Persona 1: 'The Time-Starved Executive' — Maya, 34", 
         "VP of Product at a tech scale-up; works 60 hours/week; experiences desk-bound postural strain and high cortisol.", 
         "Lacks 90 consecutive minutes for gym commutes; requires high-yield 20-minute kettlebell/bodyweight protocols that deliver metabolic conditioning and mental clarity without injury risk.", 
         "LinkedIn thought leadership, short-form YouTube Shorts during lunch breaks, Sunday evening meal prep carousels."),
        
        ("Persona 2: 'The Post-Injury Rebuilder' — David, 42", 
         "Former collegiate athlete dealing with chronic lumbar disc bulge and knee patellofemoral pain; fearful of re-injury.", 
         "Overwhelmed by aggressive cross-training workouts that exacerbate joint inflammation; seeks rigorous kinesiological form breakdowns, mobility flows, and joint longevity science.", 
         "Long-form Substack breakdowns, deep-dive YouTube mobility tutorials, in-app audio coach guidance."),
        
        ("Persona 3: 'The Consistency Novice' — Rohan, 24", 
         "Recent university graduate starting first corporate role; intimidated by gym culture and contradictory diet fads.", 
         "Experiences paralysis by analysis; struggles with motivation after missing workouts; needs frictionless habit scaffolding and non-judgmental community support.", 
         "TikTok relatable humor reels, Instagram 30-day beginner challenges, motivational push notifications.")
    ]

    for p_name, p_bg, p_need, p_chan in personas:
        add_heading_2(doc, p_name)
        log(f"### {p_name}\n")
        add_bullet_point(doc, "Demographic Profile:", p_bg)
        add_bullet_point(doc, "Core Friction & Unmet Need:", p_need)
        add_bullet_point(doc, "Preferred Content Channels:", p_chan)
        log(f"- **Demographic Profile:** {p_bg}")
        log(f"- **Core Friction & Need:** {p_need}")
        log(f"- **Preferred Channels:** {p_chan}\n")

    # Section 3: The 4 Strategic Content Pillars
    add_heading_1(doc, "3. Core Content Pillars & Tactical Topic Clusters")
    log("## 3. Core Content Pillars & Tactical Topic Clusters\n")

    p3 = ("To guarantee topical authority and content diversity, all editorial assets are categorized into four distinct content pillars. "
          "Each pillar addresses a specific stage in the customer awareness journey, balancing top-of-funnel reach with bottom-of-funnel conversion.")
    add_body_paragraph(doc, p3)
    log(p3 + "\n")

    # Pillar 1
    add_heading_2(doc, "Pillar 1: Biomechanical Mastery & Injury Prevention (30% Weight)")
    log("### Pillar 1: Biomechanical Mastery & Injury Prevention (30% Weight)\n")
    p_p1 = ("*Objective:* Demystify complex kinesiology through visual, bite-sized form corrections. Demonstrates the technical sophistication "
            "of ZenithFit's computer-vision engine while rescuing users from chronic joint pain.")
    add_body_paragraph(doc, p_p1)
    log(p_p1 + "\n")

    p1_topics = [
        ("Topic 1.1: 'The Desk Worker's Posterior Chain Fix'", "Instagram Carousel + Reel", "5 corrective drills to reverse rounded shoulders, anterior pelvic tilt, and glute amnesia caused by 8 hours of sitting."),
        ("Topic 1.2: 'Squat Depth vs. Hip Anatomy: Why One Size Does Not Fit All'", "YouTube Long-form (8 mins)", "Exploration of femoral acetabular impingement and how to adjust stance width based on individual pelvic bone morphology."),
        ("Topic 1.3: 'Stop Doing Crunches: 4 Anti-Extension Core Exercises for Bulletproof Lumbar Spines'", "TikTok / Shorts", "Visual demonstration of deadbugs, Pallof presses, and bird-dogs showing spine stabilization under load."),
        ("Topic 1.4: 'The 3-Minute Pre-Run Dynamic Warmup That Prevents Shin Splints'", "Saveable Infographic", "Ankle dorsiflexion, calf eccentric loading, and tibialis anterior activations."),
        ("Topic 1.5: 'Computer-Vision Spotlight: How ZenithFit Catches Knee Valgus in Real Time'", "Behind-the-Tech Feature", "Demonstration of our on-device optical joint-angle detection saving lifters from ACL strain.")
    ]
    for t_title, t_fmt, t_desc in p1_topics:
        add_bullet_point(doc, f"{t_title} [{t_fmt}]:", t_desc)
        log(f"- **{t_title}** *[{t_fmt}]*: {t_desc}")

    # Pillar 2
    add_heading_2(doc, "Pillar 2: Metabolic Nutrition & Pragmatic Fueling (25% Weight)")
    log("\n### Pillar 2: Metabolic Nutrition & Pragmatic Fueling (25% Weight)\n")
    p_p2 = ("*Objective:* Strip away ideological dietary dogmatism. Educates users on metabolic flexibility, bioavailable macronutrients, "
            "and sustainable cooking systems that support athletic recovery without extreme deprivation.")
    add_body_paragraph(doc, p_p2)
    log(p_p2 + "\n")

    p2_topics = [
        ("Topic 2.1: 'The 30-Gram Protein Rule: Why Breakfast Dictates Muscle Protein Synthesis'", "Carousel Guide", "Scientific breakdown of leucine thresholds for anabolic signaling and 4 quick 30g whole-food breakfast recipes."),
        ("Topic 2.2: 'Electrolytes Beyond Sugar: Why Water Alone Won't Hydrate Your Muscles'", "Short Video", "Explaining sodium-potassium ATPase pump function and simple Himalayan salt + citrus DIY hydration mixes."),
        ("Topic 2.3: 'Sunday Meal Prep for Exhausted Humans: 5 High-Protein Dinners in 45 Minutes'", "YouTube Video + Recipe Card", "Batch-cooking methodology utilizing slow-cookers and air-fryers for busy professionals."),
        ("Topic 2.4: 'Alcohol, Sleep Architecture, and Muscle Recovery: What Happens After 2 Drinks'", "Data Infographic", "Visualizing REM sleep suppression and fractional synthetic rate reductions after alcohol consumption."),
        ("Topic 2.5: 'Navigating Dining Out: A Fitness Coach's Cheat Sheet for Indian & Italian Menus'", "Substack Guide", "How to optimize protein and fiber while dining out with colleagues without social anxiety.")
    ]
    for t_title, t_fmt, t_desc in p2_topics:
        add_bullet_point(doc, f"{t_title} [{t_fmt}]:", t_desc)
        log(f"- **{t_title}** *[{t_fmt}]*: {t_desc}")

    # Pillar 3
    add_heading_2(doc, "Pillar 3: Behavioral Psychology & Habit Stacking (25% Weight)")
    log("\n### Pillar 3: Behavioral Psychology & Habit Stacking (25% Weight)\n")
    p_p3 = ("*Objective:* Address the true root cause of fitness failure—not lack of knowledge, but cognitive friction. Grounds training in "
            "James Clear's Atomic Habits and neurobiology to make movement automatic and joyful.")
    add_body_paragraph(doc, p_p3)
    log(p_p3 + "\n")

    p3_topics = [
        ("Topic 3.1: 'The 2-Minute Movement Rule: How to Beat Workout Resistance on Bad Days'", "Reel / Short", "Lowering activation energy: putting on your training shoes and completing one set as a psychological victory."),
        ("Topic 3.2: 'Dopamine Detachment: Why Motivation Is a Scam and Frictionless Rituals Win'", "Long-form Essay", "Neurological exploration of dopamine baselines and how habit stacking cements lifelong exercise adherence."),
        ("Topic 3.3: 'HRV Explained: Why Your App Told You to Rest Today (And Why You Must Listen)'", "Infographic Carousel", "De-stigmatizing rest days by showing how parasympathetic recovery drives muscle hypertrophy."),
        ("Topic 3.4: 'Overcoming the Mid-Year Fitness Slump: Reframing Goals From Scale Weight to Capacity'", "In-App Audio Reflection", "Cognitive reframing shifting focus from vanity metrics to VO2 max and deadlift power."),
        ("Topic 3.5: 'Sleep as a Performance Enhancer: The 10-3-2-1-0 Protocol'", "Visual Checklist", "Optimizing sleep hygiene: cutoff times for caffeine, food, liquid, screens, and morning alarm consistency.")
    ]
    for t_title, t_fmt, t_desc in p3_topics:
        add_bullet_point(doc, f"{t_title} [{t_fmt}]:", t_desc)
        log(f"- **{t_title}** *[{t_fmt}]*: {t_desc}")

    # Pillar 4
    add_heading_2(doc, "Pillar 4: Proof of Transformation & Community Accountability (20% Weight)")
    log("\n### Pillar 4: Proof of Transformation & Community Accountability (20% Weight)\n")
    p_p4 = ("*Objective:* Generate emotional resonance, social proof, and viral word-of-mouth through authentic user narratives, peer milestones, "
            "and seasonal interactive movement challenges.")
    add_body_paragraph(doc, p_p4)
    log(p_p4 + "\n")

    p4_topics = [
        ("Topic 4.1: 'From Chronic Back Spasms to First Pull-Up: David's 6-Month Rebuild'", "Mini-Documentary Video", "Cinematic customer profile capturing the emotional freedom of moving pain-free after years of medical frustration."),
        ("Topic 4.2: 'The ZenithFit 28-Day Mobility Reset: Community Challenge'", "Interactive App Challenge", "Global user challenge with daily 10-minute hip and thoracic spine flows, shared leaderboards, and completion badges."),
        ("Topic 4.3: 'Live Biomechanics Coach AMA: Roast My Squat Form'", "Live Stream / Stories Q&A", "Head kinesiologist reviews user-submitted video tags, correcting hip hinge and knee tracking live on camera."),
        ("Topic 4.4: 'Community Data Pulse: 1 Million Minutes of Movement Logged in October'", "Milestone Celebration Infographic", "Celebratory breakdown of collective community achievements, calories expended, and postural milestones."),
        ("Topic 4.5: 'Micro-Interviews: What Does Moving Freely Mean to You?'", "Street-Style Video Series", "Raw, unfiltered street interviews with everyday athletes discussing their personal definitions of vitality.")
    ]
    for t_title, t_fmt, t_desc in p4_topics:
        add_bullet_point(doc, f"{t_title} [{t_fmt}]:", t_desc)
        log(f"- **{t_title}** *[{t_fmt}]*: {t_desc}")

    # Section 4: Repurposing Engine
    add_heading_1(doc, "4. The 1-to-8 Content Repurposing Engine")
    log("\n## 4. The 1-to-8 Content Repurposing Engine\n")

    p4 = ("To achieve high content output without creator burnout, ZenithFit implements a hub-and-spoke production model. "
          "Every week, our head kinesiologist and performance coach record **1 Core Anchor Asset** (a 15-minute deep-dive YouTube video or podcast). "
          "From this single anchor asset, our media team systematically extracts 8 derivative micro-assets across all channels:")
    add_body_paragraph(doc, p4)
    log(p4 + "\n")

    repurpose_steps = [
        ("Anchor Asset:", "15-minute high-definition YouTube video: 'The Definitive Guide to Knee Longevity & Squat Biomechanics'"),
        ("Micro-Asset 1 & 2 (Instagram / TikTok):", "Two 45-second high-energy vertical reels highlighting specific form mistakes (knee caving vs hip drive) with on-screen kinetic graphics."),
        ("Micro-Asset 3 (Instagram Carousel):", "A 7-slide educational visual summary breaking down the 3 pre-squat mobility drills with step-by-step cue cards."),
        ("Micro-Asset 4 (LinkedIn Article):", "An executive-level thought piece connecting physical ergonomic health with C-suite cognitive performance and corporate healthcare savings."),
        ("Micro-Asset 5 (Substack Newsletter):", "A deep-dive text dispatch delivered to 85,000 subscribers containing peer-reviewed study links, nutrition pairings, and coach commentary."),
        ("Micro-Asset 6 & 7 (Twitter/X Thread & Quote Graphics):", "A 6-tweet bite-sized breakdown of joint cartilage physiology paired with high-contrast pull-quotes."),
        ("Micro-Asset 8 (In-App Push & Audio Cue):", "A 90-second audio coaching tip integrated directly into the ZenithFit app warm-up flow for that day's scheduled leg workout.")
    ]
    for r_title, r_desc in repurpose_steps:
        add_bullet_point(doc, r_title, r_desc)
        log(f"- **{r_title}** {r_desc}")

    # Section 5: Distribution & KPI Scorecard
    add_heading_1(doc, "5. Omnichannel Channel Matrix & KPI Growth Scorecard")
    log("\n## 5. Omnichannel Channel Matrix & KPI Growth Scorecard\n")

    p5 = ("Content performance is rigorously monitored through a balanced scorecard that unifies creative engagement with commercial SaaS pipeline metrics:")
    add_body_paragraph(doc, p5)
    log(p5 + "\n")

    kpis = [
        ("Organic App Install Velocity:", "Attributed downloads via UTM tracking links embedded in YouTube descriptions, bio links, and newsletter CTA buttons (Target: +25% MoM)."),
        ("Content Engagement Rate (CER):", "Average weighted engagement (comments + saves + shares / impressions) maintained above 5.2% on Instagram and 8.0% on TikTok."),
        ("High-Intent Save Rate:", "Educational carousels and recipe checklists must achieve > 10% save-to-impression ratio, confirming enduring reference utility."),
        ("Free-to-Paid Trial Conversion Rate:", "Users acquired through organic educational content exhibit a benchmark trial-to-paid subscription conversion rate of > 38% (vs 22% on paid ads)."),
        ("Day-30 & Day-90 Retention:", "Subscribers engaged with community challenges and in-app audio tips demonstrate a 90-day active retention rate exceeding 64%.")
    ]
    for k_title, k_desc in kpis:
        add_bullet_point(doc, k_title, k_desc)
        log(f"- **{k_title}** {k_desc}")

    # Section 6: Editorial Production Workflow & Operational Cadence
    add_heading_1(doc, "6. Editorial Production Workflow & Operational Cadence")
    log("\n## 6. Editorial Production Workflow & Operational Cadence\n")

    p6 = ("To maintain editorial rigor and prevent bottlenecks across creative teams, ZenithFit executes content development through "
          "a synchronized two-week sprint methodology:")
    add_body_paragraph(doc, p6)
    log(p6 + "\n")

    workflows = [
        ("Sprint Stage 1 — Ideation & Kinesiology Review (Days 1–2):", 
         "Editorial leads collaborate with registered physical therapists and nutritionists to review emerging biomechanical studies, audit app search query trends, and validate movement cues against medical safety standards."),
        ("Sprint Stage 2 — Batch Production & Filming (Days 3–5):", 
         "Filming occurs in monthly studio blocks. High-definition 4K multicam setups record core exercises from frontal and sagittal planes, capturing micro-adjustments in joint angles and foot grounding."),
        ("Sprint Stage 3 — Post-Production & Kinetic Graphics (Days 6–8):", 
         "Editors apply on-screen vector motion paths, angle callouts (e.g. 'Maintain 45° Torso Lean'), and high-contrast captioning optimized for sound-off mobile feed consumption."),
        ("Sprint Stage 4 — Quality Assurance & Pre-Flight Scheduling (Days 9–10):", 
         "All assets undergo multi-tier review in Asana and Notion: verifying anatomical terminology accuracy, testing link tracking parameters, and pre-scheduling through native platform schedulers."),
        ("Sprint Stage 5 — Real-Time Community Engagement & Loop Feedback (Ongoing):", 
         "Community managers actively moderate comment sections within the first 60 minutes of publication, directly addressing subscriber movement queries and logging recurring pain points for subsequent sprint ideation.")
    ]
    for w_title, w_desc in workflows:
        add_bullet_point(doc, w_title, w_desc)
        log(f"- **{w_title}** {w_desc}")

    docx_path = os.path.join(DOCX_DIR, "Week_05_Content_Strategy_Development_Fitness_App.docx")
    doc.save(docx_path)

    md_path = os.path.join(MD_DIR, "Week_05_Content_Strategy_Development_Fitness_App.md")
    with open(md_path, "w", encoding="utf-8") as f:
        f.write("\n".join(md_lines))

    # Calculate word count
    full_text = " ".join([p.text for p in doc.paragraphs])
    word_count = len(full_text.split())
    print(f"Week 5 Task built successfully: {docx_path} ({word_count} words)")
    return word_count

if __name__ == "__main__":
    build_week_05()
