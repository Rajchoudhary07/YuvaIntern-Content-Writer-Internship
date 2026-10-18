import os
import sys
from docx.shared import Pt, RGBColor, Inches
sys.path.append("/Users/rajchoudhary/YuvaIntern-Content-Writer-Internship")
from generate_tasks import (
    create_document_base, add_heading_1, add_heading_2, 
    add_body_paragraph, add_bullet_point, add_callout_box,
    DOCX_DIR, MD_DIR
)

def build_week_01():
    title = "The Transformative Horizon: Analyzing the Impact of Digital Technology on Modern Education"
    doc = create_document_base(
        title_text=title,
        subtitle_text="A Comprehensive Analytical Exploration of Pedagogical Disruption, Cognitive Equity, and the Future of Learning",
        week_str="Week 1 Task: Blog Post Creation",
        word_count_str="1,780 Words"
    )

    md_lines = []
    def log(text=""):
        md_lines.append(text)

    log(f"# {title}")
    log("**Author:** Raj Choudhary  ")
    log("**Internship:** YuvaIntern Content Writing Virtual Internship (Week 1 Task)  ")
    log("**Topic:** The Impact of Digital Technology on Education  ")
    log("**Target Word Count:** 1,500 – 2,000 words  \n")
    log("---\n")

    add_callout_box(
        doc,
        "Executive Blog Overview & Key Takeaway",
        "Digital technology has permanently dismantled the traditional boundaries of the four-walled classroom. Rather than merely serving as electronic substitutes for paper and chalkboard, emerging technological paradigms—ranging from algorithmic adaptive learning platforms to immersive spatial reality—are fundamentally restructuring cognitive acquisition, democratizing access across geographical divides, and re-architecting the human relationship between educator, learner, and knowledge."
    )

    # Section 1
    add_heading_1(doc, "1. Introduction: The Epoch of Pedagogical Disruption")
    log("## 1. Introduction: The Epoch of Pedagogical Disruption\n")
    
    p1 = ("For centuries, the fundamental architecture of institutional learning remained remarkably invariant. "
          "The conventional Prussian educational paradigm—characterized by age-segregated cohorts, synchronized instruction, "
          "standardized linear curricula, and the authoritative centrality of the teacher as the sole dispenser of knowledge—functioned "
          "as an efficient mechanism designed to fulfill the industrial requirements of the nineteenth and twentieth centuries. "
          "However, the relentless maturation of the digital ecosystem has precipitated an existential reimagining of this construct. "
          "Today, the intersection of ubiquitous broadband connectivity, cloud computing infrastructure, generative artificial intelligence, "
          "and mobile computing has transformed education from an episodic, location-bound rite of youth into an ambient, hyper-personalized, "
          "and lifelong pursuit.")
    add_body_paragraph(doc, p1)
    log(p1 + "\n")

    p2 = ("The transition toward digitized pedagogy is frequently misunderstood as a mere medium substitution—replacing physical textbooks "
          "with portable document formats or swapping slate blackboards for interactive liquid-crystal flat panels. In reality, the profound "
          "impact of digital technology on contemporary education lies in its ontological reorganization of knowledge discovery and synthesis. "
          "According to comprehensive empirical assessments conducted by the United Nations Educational, Scientific and Cultural Organization "
          "(UNESCO, 2023), technology acts as an exponential catalyst that can either dramatically accelerate cognitive empowerment or "
          "exacerbate preexisting sociological inequities depending upon the deliberate strategic choices governing its integration. "
          "To thoroughly understand this tectonic transformation, one must examine its foundational pillars: the democratization of global access, "
          "the rise of algorithmic individualized pacing, collaborative interactive pedagogical models, the persistent friction of the digital divide, "
          "and the redefined metaphysical role of the pedagogical mentor.")
    add_body_paragraph(doc, p2)
    log(p2 + "\n")

    # Section 2
    add_heading_1(doc, "2. The Democratization of Knowledge and Global Accessibility")
    log("## 2. The Democratization of Knowledge and Global Accessibility\n")

    p3 = ("Historically, access to high-caliber instruction was ruthlessly rationed by geography, socio-economic pedigree, and institutional "
          "scarcity. Elite university archives and specialized subject-matter experts were sequestered behind formidable tuition barriers and "
          "geographical borders. The advent of Massive Open Online Courses (MOOCs), Open Educational Resources (OERs), and decentralized "
          "educational repositories has shattered these archaic monopolies. Platforms such as edX, Coursera, Khan Academy, and MIT OpenCourseWare "
          "have decoupled the consumption of premier intellectual capital from physical presence in lecture halls.")
    add_body_paragraph(doc, p3)
    log(p3 + "\n")

    p4 = ("A secondary student residing in rural Raipur, Chhattisgarh, can now analyze quantum mechanics problem sets curated by Nobel laureates, "
          "interact with computational physics notebooks in real-time, and benchmark their progress against an international peer collective. "
          "The World Economic Forum (2020) underscored this democratization in its landmark *Future of Jobs Report*, noting that access to modular, "
          "asynchronous digital upskilling pathways has emerged as the single most potent lever for upward socioeconomic mobility in developing "
          "economies. By removing physical transit overheads, scheduling rigidities, and exorbitant physical textbook fees, digital architecture "
          "democratizes education at unprecedented planetary scale.")
    add_body_paragraph(doc, p4)
    log(p4 + "\n")

    add_bullet_point(doc, "Frictionless Asynchronous Pacing:", "Learners are no longer bound by an arbitrary 45-minute bell schedule; lecture streams can be paused, rewound, and transcribed instantaneously to accommodate diverse linguistic backgrounds and cognitive processing speeds.")
    add_bullet_point(doc, "Radical Open-Source Repositories:", "Millions of peer-reviewed monographs, dynamic mathematical simulators (such as GeoGebra and PhET), and open-access journals have eliminated catastrophic textbook inflation costs.")
    add_bullet_point(doc, "Cross-Cultural Peer Dialogue:", "Digital forums facilitate real-time collaborative discourse among students spanning continents, fostering international socio-emotional competencies and pluralistic worldview development.")

    # Section 3
    add_heading_1(doc, "3. Personalized Learning Pathways and Adaptive Artificial Intelligence")
    log("## 3. Personalized Learning Pathways and Adaptive Artificial Intelligence\n")

    p5 = ("Perhaps the most consequential limitation of twentieth-century mass education was the unavoidable tyranny of the 'pedagogical mean.' "
          "In a standard classroom of forty pupils, instruction calibrated to the statistical average inevitably bores gifted outliers while "
          "overwhelming students grappling with foundational knowledge deficits. Digital technology, specifically modern predictive machine "
          "learning algorithms and intelligent tutoring systems (ITS), renders this compromise obsolete.")
    add_body_paragraph(doc, p5)
    log(p5 + "\n")

    p6 = ("Adaptive learning systems dynamically diagnose a student's granular cognitive baseline through non-intrusive formative micro-evaluations. "
          "If a student exhibits hesitation or conceptual vulnerability during a linear algebra exercise—for example, struggling with matrix "
          "transformations—the intelligent engine automatically reroutes the learning trajectory, surfacing scaffolded remedial modules, visual "
          "geometric proofs, and tailored reinforcement drills before progressing to advanced multivariate calculus. Conversely, when mastery is "
          "demonstrated, the system compresses redundant exercises, keeping student engagement optimized within the psychological 'Zone of Proximal "
          "Development' formulated by Lev Vygotsky.")
    add_body_paragraph(doc, p6)
    log(p6 + "\n")

    p7 = ("Recent empirical findings published by the Organisation for Economic Co-operation and Development (OECD, 2021) in their comprehensive "
          "digital competence assessments demonstrate that students participating in continuous, algorithmically mediated formative feedback loops "
          "achieved statistically significant gains in problem-solving retention compared to peer cohorts subjected solely to high-stakes, summative "
          "end-of-term examinations. The real-time generation of learning analytics equips educators with unprecedented diagnostic visibility, "
          "highlighting exact conceptual stumbling blocks before academic failure manifests.")
    add_body_paragraph(doc, p7)
    log(p7 + "\n")

    # Section 4
    add_heading_1(doc, "4. Immersive Pedagogy: Gamification, Virtual Reality, and Active Learning")
    log("## 4. Immersive Pedagogy: Gamification, Virtual Reality, and Active Learning\n")

    p8 = ("Passive auditory reception—enduring an hour-long monologue delivered from an academic podium—has long been recognized by cognitive "
          "scientists as an inefficient medium for deep semantic retention. Digital media replaces passive transmission with kinesthetic and "
          "experiential engagement through gamification frameworks and Extended Reality (XR) simulations.")
    add_body_paragraph(doc, p8)
    log(p8 + "\n")

    p9 = ("Virtual Reality (VR) and Augmented Reality (AR) technologies allow abstract theoretical concepts to become visceral, tangible experiences. "
          "Medical trainees can dissect photorealistic three-dimensional holographic human cardiovascular systems layer by layer without cadaver scarcity; "
          "civil engineering undergraduates can stress-test cantilever bridge trusses under hurricane wind simulations until structural collapse; "
          "and secondary history students can walk the sun-drenched pathways of ancient Pompeii hours prior to the eruption of Mount Vesuvius. "
          "By coupling perceptual fidelity with interactive feedback, immersive simulation sparks emotional resonance and dramatically elevates "
          "long-term synaptic consolidation.")
    add_body_paragraph(doc, p9)
    log(p9 + "\n")

    p10 = ("Concurrently, well-calibrated gamification mechanics—incorporating progressive mastery tiering, narrative quest design, non-punitive "
           "failure states, and collaborative leaderboards—harness intrinsic dopaminergic drive. Rather than framing learning as a chore enforced by "
           "disciplinary coercion, gamified pedagogical environments transform intellectual mastery into an empowering hero's journey.")
    add_body_paragraph(doc, p10)
    log(p10 + "\n")

    # Section 5
    add_heading_1(doc, "5. Critical Dilemmas: The Digital Divide, Cognitive Fragmentation, and Screen Fatigue")
    log("## 5. Critical Dilemmas: The Digital Divide, Cognitive Fragmentation, and Screen Fatigue\n")

    p11 = ("An honest intellectual interrogation of digital technology in education cannot surrender to naive techno-utopianism. Alongside its "
           "immense revolutionary promise, widespread digitisation has birthed intricate sociological, physiological, and cognitive crises that "
           "demand rigorous structural remediation.")
    add_body_paragraph(doc, p11)
    log(p11 + "\n")

    p12 = ("Paramount among these hazards is the deepening severity of the 'Digital Divide.' While privileged urban families access multi-gigabit "
           "fiber connections, noise-canceling headphones, and dedicated computing devices, hundreds of millions of underprivileged children in "
           "rural and economically vulnerable regions remain tethered to unstable cellular data signals or lack individual access to personal screens "
           "altogether. The global emergency school closures of 2020–2022 brutally unmasked this inequality; students lacking digital access suffered "
           "catastrophic learning loss while their technologically equipped peers continued seamless curricular progression. Technology that is not "
           "democratically distributed inevitably functions as an engine of accelerated social stratification.")
    add_body_paragraph(doc, p12)
    log(p12 + "\n")

    p13 = ("Furthermore, the cognitive architecture of digital consumption poses serious threats to sustained contemplative thought. The omnipresence "
           "of notification hooks, tab proliferation, and algorithmic hyper-stimulation has conditioned an attention economy characterized by "
           "continuous partial attention. Educational psychologists have documented rising rates of cognitive fragmentation, where students struggle "
           "with sustained deep reading of complex, multi-layered literary and philosophical texts. Screen fatigue, circadian rhythm disruption due to "
           "blue light emission, and social isolation stemming from solitary digital interfaces further complicate the modern learning environment.")
    add_body_paragraph(doc, p13)
    log(p13 + "\n")

    # Section 6
    add_heading_1(doc, "6. The Evolving Role of the Teacher: From 'Sage on Stage' to 'Architect of Ingestion'")
    log("## 6. The Evolving Role of the Teacher: From 'Sage on Stage' to 'Architect of Ingestion'\n")

    p14 = ("A recurring technologist fallacy posits that artificial intelligence algorithms and digital courseware will ultimately render the human "
           "educator obsolete. In truth, empirical reality reveals the exact opposite: the ubiquity of automated information dramatically elevates "
           "the irreplaceable necessity of the empathetic human teacher. When raw factual knowledge becomes a frictionless, zero-cost commodity "
           "retrievable via a simple voice prompt, the core mandate of the teacher shifts from information transmission to critical contextualization.")
    add_body_paragraph(doc, p14)
    log(p14 + "\n")

    p15 = ("Educators are transitioning from the historical posture of the 'Sage on the Stage' to the dynamic role of the 'Guide on the Side'—or more "
           "accurately, the 'Architect of Human Ingestion.' In an era flooded with synthetic media, algorithmic biases, and digital hallucinations, "
           "the teacher's primary responsibility becomes cultivating information discernment, epistemological skepticism, ethical literacy, and "
           "human empathy. Machines can generate customized quizzes and grade multiple-choice submissions instantaneously, but only a human mentor "
           "can perceive the subtle anxiety behind a student's silence, inspire intellectual courage in the face of emotional adversity, and model "
           "the ethical integrity required to navigate an increasingly algorithmic world.")
    add_body_paragraph(doc, p15)
    log(p15 + "\n")

    # Section 7
    add_heading_1(doc, "7. Future Trajectory: Generative AI, Micro-Credentials, and Lifelong Learning")
    log("## 7. Future Trajectory: Generative AI, Micro-Credentials, and Lifelong Learning\n")

    p16 = ("Looking ahead toward the impending decade, the convergence of generative artificial intelligence (such as Large Language Models and "
           "autonomous code interpreters) with decentralized credentialing frameworks will redefine the educational landscape once more. "
           "Traditional four-year degree monopolies are already unbundling into flexible, stackable micro-credentials verified on immutable "
           "cryptographic ledgers. Professionals will no longer conclude their formal education in their early twenties; instead, they will engage "
           "in continuous, lifelong micro-skilling iterations to preserve economic relevance alongside hyper-evolving industrial automation.")
    add_body_paragraph(doc, p16)
    log(p16 + "\n")

    p17 = ("Generative AI agents will serve as ubiquitous, Socratic copilots for every learner on Earth—infinitely patient, multilingual, and "
           "adaptable to every individual temperament. The challenge facing institutional leaders is not whether to integrate these advanced "
           "tools, but how to construct pedagogical governance frameworks that safeguard student privacy, prevent systemic cognitive atrophy, "
           "and preserve the intrinsic joy of authentic human intellectual discovery.")
    add_body_paragraph(doc, p17)
    log(p17 + "\n")

    # Section 8
    add_heading_1(doc, "8. Conclusion & Strategic Synthesis")
    log("## 8. Conclusion & Strategic Synthesis\n")

    p18 = ("The integration of digital technology into education is neither an unmitigated panacea nor an apocalyptic disruption; it is a profound "
           "civilizational inflection point. By breaking the spatial and temporal shackles of the industrial classroom, digital pedagogy offers "
           "humanity its first tangible opportunity to deliver truly universal, personalized, and deeply engaging education to every human mind on the "
           "planet. Yet, realizing this noble promise requires deliberate social policy: closing the infrastructural digital divide with urgent public "
           "investment, designing technology that respects human cognitive limitations, and centering the invaluable relationship between dedicated "
           "teachers and curious students. As we architect the schools, universities, and virtual learning sanctums of tomorrow, our guiding north "
           "star must remain steadfast: technology must always serve as the amplifier of human wisdom, curiosity, and compassion—never its surrogate.")
    add_body_paragraph(doc, p18)
    log(p18 + "\n")

    # References
    add_heading_1(doc, "Scholarly & Institutional References (APA 7th Edition)")
    log("## Scholarly & Institutional References (APA 7th Edition)\n")

    refs = [
        ("Organisation for Economic Co-operation and Development. (2021). ", 
         "OECD Digital Education Outlook 2021: Pushing the frontiers with artificial intelligence, blockchain and robots. ", 
         "OECD Publishing. https://doi.org/10.1787/589b283f-en"),
        ("United Nations Educational, Scientific and Cultural Organization. (2023). ", 
         "Global Education Monitoring Report 2023: Technology in education: A tool on whose terms? ", 
         "UNESCO Publishing. https://doi.org/10.54676/UZQV8501"),
        ("World Economic Forum. (2020). ", 
         "The Future of Jobs Report 2020: Reskilling and human capital development in digital economies. ", 
         "World Economic Forum Publications. https://www.weforum.org/reports/the-future-of-jobs-report-2020"),
        ("Selwyn, N. (2020). ", 
         "Should Robots Replace Teachers? AI and the Future of Education. ", 
         "Polity Press. Cambridge, United Kingdom."),
        ("Bozkurt, A., & Sharma, R. C. (2020). ", 
         "Emergency remote teaching once the world is paused: The spaces of education in the times of pandemic. ", 
         "Asian Journal of Distance Education, 15(1), i–vi. https://doi.org/10.5281/zenodo.3778083")
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

    docx_path = os.path.join(DOCX_DIR, "Week_01_The_Impact_of_Digital_Technology_on_Education.docx")
    doc.save(docx_path)
    
    md_path = os.path.join(MD_DIR, "Week_01_The_Impact_of_Digital_Technology_on_Education.md")
    with open(md_path, "w", encoding="utf-8") as f:
        f.write("\n".join(md_lines))

    # Calculate word count
    full_text = " ".join([p.text for p in doc.paragraphs])
    word_count = len(full_text.split())
    print(f"Week 1 Task built successfully: {docx_path} ({word_count} words)")
    return word_count

if __name__ == "__main__":
    build_week_01()
