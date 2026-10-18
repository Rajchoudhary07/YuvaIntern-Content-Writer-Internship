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

def build_week_02():
    title = "TerraThreads: 30-Day Comprehensive Social Media Content Strategy & Editorial Calendar"
    doc = create_document_base(
        title_text=title,
        subtitle_text="A Month-Long Omnichannel Content Blueprint for a Circular, Ethical & Regenerative Fashion Label",
        week_str="Week 2 Task: Social Media Content Planning",
        word_count_str="30 Daily Posts + Brand Architecture (2,450 Words)"
    )

    md_lines = []
    def log(text=""):
        md_lines.append(text)

    log(f"# {title}")
    log("**Brand:** TerraThreads (Circular & Ethical Fashion Label)  ")
    log("**Author:** Raj Choudhary  ")
    log("**Internship:** YuvaIntern Content Writing Virtual Internship (Week 2 Task)  ")
    log("**Deliverable:** 30 Distinct Post Plans with Dates, Pillars, Formats, Full Captions, CTAs, and Visual Directions  \n")
    log("---\n")

    add_callout_box(
        doc,
        "Strategic Campaign Objective",
        "This editorial content plan establishes TerraThreads as the premier thought leader in circular, regenerative fashion. Over 30 consecutive days, the strategy balances organic brand storytelling, radical supply chain transparency, wearable capsule styling, and interactive eco-education to drive community engagement, cultivate consumer trust, and convert mindful browsers into lifelong brand advocates."
    )

    # Brand Architecture
    add_heading_1(doc, "1. Brand Identity & Strategic Foundation: TerraThreads")
    log("## 1. Brand Identity & Strategic Foundation: TerraThreads\n")

    p1 = ("TerraThreads is an ethical, direct-to-consumer apparel brand founded on the principles of regenerative agriculture, "
          "closed-loop textile manufacturing, and living-wage artisan craftsmanship. Unlike fast-fashion retailers operating on rapid "
          "obsolescence and resource depletion, TerraThreads designs timeless, modular wardrobe essentials woven exclusively from "
          "GOTS-certified organic rain-fed cotton, European flax linen, and botanical natural dyes derived from upcycled agricultural waste.")
    add_body_paragraph(doc, p1)
    log(p1 + "\n")

    add_bullet_point(doc, "Brand Mission:", "To eliminate textile waste through circular design, radical transparency, and heirloom-quality craftsmanship that honours both the earth and the artisan.")
    add_bullet_point(doc, "Target Audience:", "Eco-conscious millennials and Gen Z professionals (aged 22–38), urban minimalists, conscious consumers, and sustainable living advocates who prioritize garment longevity and ethical provenance over fleeting micro-trends.")
    add_bullet_point(doc, "Brand Voice & Persona:", "Mindful, transparent, elevated, grounded, and empowering. Educational without being pedantic; inspiring without greenwashing.")
    add_bullet_point(doc, "Primary Platforms:", "Instagram (Aesthetic visual storytelling, Carousels & Reels), TikTok / YouTube Shorts (Raw behind-the-seams & dye tutorials), LinkedIn (B2B supply chain ethics and ESG thought leadership), Pinterest (Capsule wardrobe moodboards).")

    # Content Pillars
    add_heading_1(doc, "2. Core Strategic Content Pillars")
    log("## 2. Core Strategic Content Pillars\n")

    pillars = [
        ("Pillar 1: Radical Transparency & Supply Chain (25%)", 
         "Deconstructing the 'cost-per-wear', introducing our certified fair-trade weaving cooperatives in India, demonstrating plant-based dye alchemy, and publishing open-source carbon balance audits."),
        ("Pillar 2: Capsule Styling & Versatility (30%)", 
         "Demonstrating how 7 timeless modular pieces unlock 30+ distinctive looks across workwear, weekend leisure, and evening events; highlighting wardrobe longevity and anti-consumerist curation."),
        ("Pillar 3: Eco-Education & Myth-Busting (25%)", 
         "Unpacking complex textile science in bite-sized infographics: microplastics in synthetic polyester, greenwashing terminology, water footprint of conventional denim, and garment care preservation."),
        ("Pillar 4: Community, Artisans & Lifestyle Rituals (20%)", 
         "Celebrating our artisan masters, sharing real customer wear stories, hosting mindful living rituals, and initiating interactive circularity challenges (e.g., #30WearRule).")
    ]
    for p_title, p_desc in pillars:
        add_bullet_point(doc, p_title + ":", p_desc)
        log(f"- **{p_title}:** {p_desc}")

    # 30-Day Master Calendar Table
    add_heading_1(doc, "3. The 30-Day Master Editorial Calendar")
    log("\n## 3. The 30-Day Master Editorial Calendar\n")

    posts = [
        # Week 1
        ("Day 01 (Oct 01)", "Instagram Carousel", "Pillar 1: Transparency", 
         "The Anatomy of an Honest Shirt: What does ethical fashion really cost?",
         "Most fashion labels hide behind opaque markups. We believe you deserve the raw truth. Swipe through to see the exact breakdown of our Classic Linen Overshirt: from rain-fed flax cultivation to living-wage hand-weaving in Maheshwar, non-toxic madder-root dyeing, and compostable seed-paper tags.\n\nFair wages aren't an added perk; they are our non-negotiable foundation.\n\nSave this post for your conscious shopping checklist!",
         "Save & Share", "Clean split-screen studio photography showing garment cost breakdown infograph against raw linen textures.",
         "#EthicalFashion #RadicalTransparency #WhoMadeMyClothes #SlowFashion #TerraThreads"),
        
        ("Day 02 (Oct 02)", "Reels / TikTok", "Pillar 3: Eco-Education", 
         "Did you know 60% of modern clothing is made of hidden plastic?",
         "Check your clothing tag right now. If it reads polyester, nylon, acrylic, or polyamide—you're wearing fossil-fuel plastic. Every wash sheds up to 700,000 microplastic fibers into our waterways. At TerraThreads, our fibers grow from the soil and can return to the soil safely. What material makes up the shirt you're wearing today? Drop it below!",
         "Comment your tag material", "Fast-paced macro lens transition showing synthetic fibers shedding vs. natural organic cotton decomposing.",
         "#Microplastics #PlasticFreeFashion #SustainableLiving #TextileScience #CleanClothing"),

        ("Day 03 (Oct 03)", "Instagram Single / Pinterest", "Pillar 2: Capsule Styling", 
         "The 7-Piece Autumn Minimalist Capsule",
         "You don't need a sprawling walk-in wardrobe to look effortlessly put-together every day. You just need pieces cut with architectural precision from breathable, natural textures. Here is our 7-piece Autumn Foundation: 1 Crisp Hemp Shirt, 2 Wide-Leg Linen Trousers, 1 Organic Knit Sweater, 1 Reversible Chore Coat, and 2 Essential Cotton Tees.\n\nTap the link in bio to download our Free Modular Styling Lookbook.",
         "Click link in bio", "Warm aesthetic flat-lay on reclaimed light timber flooring with earth-toned ceramic accents.",
         "#CapsuleWardrobe #MinimalistStyle #OutfitInspiration #SlowLiving #TimelessFashion"),

        ("Day 04 (Oct 04)", "LinkedIn Article / Post", "Pillar 1: Transparency", 
         "Why Fast Fashion's Supply Chain Is an Ecological Emergency",
         "Over 100 billion garments are produced annually, and an estimated 85% end up incinerated or rotting in landfills in the Global South within 12 months of production. As apparel executives, we must confront a fundamental truth: linear fashion models are ecologically bankrupt. Here is how TerraThreads built a zero-waste closed-loop supply system from day one.",
         "Join the B2B discussion", "Infographic diagram depicting linear fast fashion waste vs. TerraThreads circular regenerative loop.",
         "#Sustainability #ESG #SupplyChainEthics #CircularEconomy #FashionLeadership"),

        ("Day 05 (Oct 05)", "Instagram Stories / Poll", "Pillar 4: Community", 
         "The #30WearChallenge: Can you commit before you click buy?",
         "Next time you're about to buy a piece of clothing, pause and ask yourself: 'Will I wear this at least 30 times?' If the answer isn't an enthusiastic yes, leave it on the rack. Vote in our poll: How many items in your current closet have you worn 30+ times? (A: Over 80% | B: About 50% | C: Working on it!)",
         "Vote in poll & DM us", "Interactive story cards with slider sticker and user response reposts.",
         "#30Wears #MindfulConsumption #ClosetAudit #TerraThreadsCommunity"),

        ("Day 06 (Oct 06)", "Reels / Short Video", "Pillar 1: Transparency", 
         "Meet Master Weaver Rajesh: 28 years preserving handloom heritage",
         "The rhythmic clack of the wooden loom is music in Maheshwar. Master artisan Rajesh has been perfecting hand-spun weaving for nearly three decades. When you choose handloom, you aren't just wearing fabric—you are preserving ancestral heritage and ensuring generational artisan dignity. Thank you for honouring the hands that weave your days.",
         "Double-tap to honour artisans", "Documentary-style warm cinematic footage of Rajesh weaving at golden hour, capturing shuttle movement and smile.",
         "#ArtisanMade #HandloomHeritage #Craftsmanship #EthicalMakers #SlowCraft"),

        ("Day 07 (Oct 07)", "Instagram Carousel", "Pillar 3: Eco-Education", 
         "Greenwashing Alert: 5 Terms Fast Fashion Brands Use to Trick You",
         "Don't let buzzwords fool you. In today's guide: 1) 'Conscious Collection' (often contains 5% recycled poly), 2) 'Eco-Friendly' (unregulated marketing jargon), 3) 'Vegan Leather' (99% PVC / toxic polyurethane plastic), 4) 'Carbon Neutral Shipping' (often unverified offsets), 5) 'Natural' (doesn't mean pesticide-free or non-toxic).\n\nBookmark this guide for your next shopping excursion!",
         "Save for reference", "Bold, typography-led carousels using contrast yellow and slate black highlight banners.",
         "#Greenwashing #ConsciousConsumer #FashionTruth #EcoAwareness #ShopEthical"),

        # Week 2
        ("Day 08 (Oct 08)", "Instagram Reel", "Pillar 2: Capsule Styling", 
         "1 Linen Dress, 5 Completely Different Settings",
         "From a Sunday morning farmers market run to boardroom client presentations and candlelight rooftop dinners—here is how intentional design does the heavy lifting. Watch how shifting footwear, belts, and layering pieces transforms our A-line Flax Smock Dress. Which look matches your personal aesthetic best?",
         "Comment Look 1, 2, 3, 4, or 5", "Seamless match-cut transition reel with upbeat acoustic jazz soundtrack.",
         "#StylingHacks #VersatileWardrobe #WearMoreSpendLess #LinenDress #EffortlessStyle"),

        ("Day 09 (Oct 09)", "TikTok / Reels", "Pillar 1: Transparency", 
         "We boil pomegranate rinds and turmeric to dye our fabrics",
         "Chemical synthetic dyes pollute up to 20% of global industrial water. At TerraThreads, our autumn ochre hues come from discarded temple marigolds and pomegranate rinds sourced from local juice stalls. Watch our head dyer Anita extract liquid sunshine. Zero toxic petrochemicals, 100% biodegradable wash-off.",
         "Drop a 🌿 if you love botanical dye", "ASMR sensory video capturing boiling dye vats, fabric immersion, and sun drying.",
         "#NaturalDyes #BotanicalColors #ZeroWasteAlchemy #PlantDyed #CircularDesign"),

        ("Day 10 (Oct 10)", "Instagram Infographic", "Pillar 3: Eco-Education", 
         "Textile Care 101: How washing your clothes less actually saves them",
         "The most sustainable garment is the one already in your closet. Frequent hot water cycles and tumble dryers degrade organic fibers and waste immense energy. Tip: 1) Wash at 30°C or cold water, 2) Air-dry in gentle shade, 3) Spot clean small stains with natural soap, 4) Air out wool and linen between wears. Double the life of your wardrobe with simple mindfulness.",
         "Share with a laundry friend", "Hand-drawn botanical washing icon diagrams with muted earthy backgrounds.",
         "#GarmentCare #SustainableLaundry #EcoTips #WardrobeLongevity #SlowLiving"),

        ("Day 11 (Oct 11)", "Instagram Stories", "Pillar 4: Community", 
         "Ask Me Anything: Live Q&A with our Founder on Ethical Sourcing",
         "Have questions about our GOTS certification, how our artisan wages compare to standard living benchmarks, or how our plant dyes hold up in the wash? Drop your burning questions into the question sticker—our founder Ananya is answering all questions live from the dyeing studio today at 5 PM IST!",
         "Drop questions in sticker", "Behind-the-scenes handheld video clips from the dye studio with Q&A sticker overlay.",
         "#FounderQnA #BehindTheBrand #RadicalTransparency #TerraThreadsLive"),

        ("Day 12 (Oct 12)", "Pinterest / Instagram", "Pillar 2: Capsule Styling", 
         "Earth Tone Color Harmony: Curating an Organic Palette",
         "Nature never clashes. When you build a wardrobe around moss greens, raw ecru, warm terracotta, slate charcoal, and indigo, every single top pairs harmoniously with every single bottom. Pin this palette for your seasonal wardrobe moodboard.",
         "Pin to Pinterest", "Color swatch board showing natural raw textiles alongside forest, clay, and river stones.",
         "#ColorPalette #EarthTones #MinimalistAesthetic #NeutralStyle #Moodboard"),

        ("Day 13 (Oct 13)", "Instagram Reel", "Pillar 3: Eco-Education", 
         "Why Conventional Cotton Is Not as 'Natural' as You Think",
         "While cotton is a natural plant, conventional non-organic cotton occupies only 2.5% of world agricultural land but consumes over 16% of global insecticide applications, devastating groundwater and smallholder farmer health. Organic rain-fed cotton uses zero synthetic pesticides and preserves soil microbiome vitality. Vote for clean soil with your daily clothing choices.",
         "Share to spread awareness", "Split-screen micro-documentary showing cracked pesticide soil vs. lush bio-diverse organic cotton fields.",
         "#OrganicCotton #SoilHealth #RegenerativeFarming #CleanFashion #TerraThreads"),

        ("Day 14 (Oct 14)", "Instagram Carousel", "Pillar 4: Community", 
         "Community Spotlight: How Architect Sneha wears TerraThreads on site & studio",
         "Sneha leads sustainable urban architecture projects across Mumbai. She needs garments that breathe during sweltering construction site visits yet project poise during investor presentations. Swipe to see how she pairs the Terra Oversized Hemp Blazer across a 14-hour workday. 'Clothing should liberate your movement, not constrain it.'",
         "Tag an inspiring woman below", "Editorial portrait series combining raw concrete architectural textures with flowing linen silhouettes.",
         "#WomenOfTerra #RealPeopleRealStyle #ArchitectStyle #ConsciousProfessionals"),

        # Week 3
        ("Day 15 (Oct 15)", "Instagram Carousel", "Pillar 1: Transparency", 
         "Our Zero-Plastic Packaging: From Seed-Paper Tags to Cornstarch Mailers",
         "Unboxing our garments should leave zero footprint on the planet. Here is the journey of our packaging: 1) Certified home-compostable cornstarch shipping mailers, 2) Unbleached hemp string ties, 3) Plantable wild-flower seed tags (soak them in soil and watch marigolds bloom!). Zero microplastics, 100% circular.",
         "Tell us: have you planted your tag?", "Step-by-step unboxing photography demonstrating planting the seed tag in a clay pot.",
         "#ZeroWastePackaging #PlasticFreeShipping #CompostablePackaging #PlantYourTags"),

        ("Day 16 (Oct 16)", "Reels / TikTok", "Pillar 2: Capsule Styling", 
         "Styling 1 White Hemp Shirt for 3 Seasons: Summer, Monsoon, Winter",
         "A great shirt has no expiration date. Watch how we style our signature Heirloom Hemp Shirt: 1) Summer: unbuttoned over swimwear as a breathable sun protector, 2) Monsoon: tucked into high-waist organic chinos with rolled sleeves, 3) Winter: layered under a chunky recycled wool knit with collar popped. One investment, 365 days of comfort.",
         "Save for seasonal styling", "Smooth rhythmic seasonal transition video with lo-fi organic beats.",
         "#SeasonlessStyle #WardrobeWorkhorse #HempClothing #AllSeasonStyle"),

        ("Day 17 (Oct 17)", "LinkedIn Thought Leadership", "Pillar 1: Transparency", 
         "Why We Pay Our Artisans 2.4x the Regional Minimum Wage",
         "Minimum wage in the textile sector is frequently an institutionalized poverty trap. At TerraThreads, all our partner weaving cooperatives operate on living-wage benchmarks verified by the Asia Floor Wage Alliance. When artisans are paid with dignity, craftsmanship flourishes, heritage survives, and local communities thrive. True sustainability begins with fair compensation.",
         "Share with ethical business leaders", "Professional data infographic comparing statutory minimum wages against calculated living costs in Madhya Pradesh.",
         "#FairWages #LaborRights #ArtisanEconomy #SocialResponsibility #ESGLeadership"),

        ("Day 18 (Oct 18)", "Instagram Single / Quote", "Pillar 3: Eco-Education", 
         "\"Buy less, choose well, make it last.\" — Vivienne Westwood",
         "In an algorithm-driven world urging you to buy daily trend-hauls, the most radical act of rebellion is taking exceptional care of what you already own. When you do purchase, invest in quality that your grandchildren could cherish. What is the oldest piece of clothing currently in your wardrobe? Share its story in the comments!",
         "Share your oldest garment story", "Minimalist, editorial serif typography on handmade textured unbleached cotton paper.",
         "#VivienneWestwood #SlowFashionQuote #MindfulLiving #EthicalPhilosophy"),

        ("Day 19 (Oct 19)", "Reels / Short Video", "Pillar 4: Community", 
         "Behind the Seams: A Day in Our Zero-Waste Design Studio",
         "Take a morning walk through our design studio in Bangalore. From morning chai to draping zero-waste cutting patterns that leave less than 1% fabric scrap, our team works with calm mindfulness. Any fabric remnant left over is shredded and repurposed into paper tags and pocket linings. Nothing wasted, everything cherished.",
         "Follow for daily ethical design", "Sensory, peaceful vlog aesthetic with natural morning sunlight and ambient fabric cutting shears sound.",
         "#StudioVlog #ZeroWasteDesign #BehindTheScenes #DayInTheLife #DesignStudio"),

        ("Day 20 (Oct 20)", "Instagram Carousel", "Pillar 3: Eco-Education", 
         "The Hidden Water Cost of a Single Conventional Pair of Jeans",
         "It takes approximately 7,500 liters of fresh water to produce one conventional cotton pair of denim jeans—the equivalent of what an average human drinks over seven years. Our organic, indigo-dipped botanical denim uses 84% less water through closed-loop rainwater recycling. Swipe to see the comparative footprint.",
         "Share to educate friends", "Compelling data visualization comparing 500 water buckets against our closed-loop bottle recycling system.",
         "#WaterConservation #DenimFootprint #WaterCrisis #TextileFacts #TerraThreads"),

        ("Day 21 (Oct 21)", "Instagram Stories / Quiz", "Pillar 3: Eco-Education", 
         "Sustainable Fashion Quiz: Can you identify the eco-myth?",
         "Test your sustainable fashion IQ with our 4-question interactive Sunday quiz: Question 1: Is bamboo viscose always eco-friendly? (Spoiler: heavy chemical processing!). Question 2: Does dry-cleaning harm garments? Tap through each story to test your knowledge and reveal the ecological facts!",
         "Tap stickers to reveal answers", "Series of 5 interactive quiz sticker stories with animated correct/incorrect reveals.",
         "#SundayQuiz #FashionKnowledge #EcoMythbusting #LearnWithTerra"),

        # Week 4
        ("Day 22 (Oct 22)", "Instagram Reel", "Pillar 2: Capsule Styling", 
         "How to Master the French Tuck & Garment Proportions",
         "A great outfit is 20% garment selection and 80% proportion styling. Watch our head stylist demonstrate: 1) The effortless loose French tuck, 2) The cuff roll for boxy sleeve balancing, 3) Belt placement for defining organic silhouettes. Simple intentional tweaks that elevate your daily mirror look in 10 seconds.",
         "Save this hack for tomorrow", "Vertical split demonstration showing 'Before vs After' subtle proportion adjustments.",
         "#StyleHacks #StylingTips #FrenchTuck #ProportionStyling #CleanLines"),

        ("Day 23 (Oct 23)", "TikTok / Reels", "Pillar 1: Transparency", 
         "Drop-Testing Our Garment Seams: Why Our Clothes Don't Rip",
         "Fast fashion seams unravel after three machine washes because manufacturers use single-chain cheap polyester threading. At TerraThreads, every stress point is reinforced with double-needle French seams and bartack stitching using high-tensile cotton thread. Watch us stress-test our linen seam strength on camera. Built to last decades, not seasons.",
         "Drop a comment if you hate unravelling hems", "Close-up macro lens video showing heavy mechanical pulling on seams, demonstrating flawless tensile durability.",
         "#QualityOverQuantity #CraftsmanshipDetails #FrenchSeams #BuiltToLast"),

        ("Day 24 (Oct 24)", "Instagram Carousel", "Pillar 4: Community", 
         "Voices of Our Community: 1 Year Later with the Terra Trench Coat",
         "We reached out to 5 customers who purchased our signature Hemp Trench Coat exactly 12 months ago to inspect how the fabric has aged. The unanimous consensus? 'It gets softer with every rain, molds to my body, and holds its silhouette flawlessly.' When garments are made with natural fibers, age is a badge of beauty.",
         "Read full reviews in bio", "Carousel featuring user-generated candid travel photos in Tokyo, London, and Manali wearing the trench.",
         "#RealReviews #AgedGracefully #CustomerStories #HempTrench #TerraCommunity"),

        ("Day 25 (Oct 25)", "LinkedIn Post", "Pillar 1: Transparency", 
         "The Circular Economy in Action: Announcing our Take-Back Program",
         "Today marks a milestone in TerraThreads' sustainability journey: the launch of 'TerraCycle'—our guaranteed circular take-back initiative. When any TerraThreads garment reaches the end of its wearable life, return it to us. We will deconstruct the fibers for mechanical recycling or compost it in our industrial soil beds. In exchange, receive a 20% lifetime repair credit.",
         "Read the whitepaper", "Sleek infographic depicting the circular cradle-to-cradle product lifecycle.",
         "#CircularEconomy #ClosedLoop #TakeBackProgram #ExtendedProducerResponsibility"),

        ("Day 26 (Oct 26)", "Instagram Single / Video", "Pillar 3: Eco-Education", 
         "The Power of Sunlight: Why Line-Drying Clothes Matters",
         "Electric tumble dryers account for roughly 6% of an average home's electricity consumption. Beyond carbon emissions, the violent tumbling cycle damages delicate natural fibers. Hang your linen and cotton out in the open morning breeze. Sunlight naturally disinfects whites, preserves fiber tension, and leaves clothing with the crisp aroma of fresh air.",
         "Double-tap if you love sun-dried linen", "Sunlight-drenched slow-motion video of crisp white linen waving gently on a breezy rooftop washing line.",
         "#LineDrying #SunDried #SlowRituals #SimpleLiving #EcoHabits"),

        ("Day 27 (Oct 27)", "Instagram Reel", "Pillar 2: Capsule Styling", 
         "Packing for a 5-Day Work Trip in Just One Carry-On Bag",
         "Never check luggage again. Here is how we pack 5 days of versatile workwear and evening leisure in a compact 35L cabin bag using lightweight, wrinkle-relaxing linen and organic cotton essentials. Watch our roll-packing technique that prevents creasing and maximizes compartment space.",
         "Save this for your next business trip", "Overhead stop-motion packing video smoothly nesting garments into a sleek recycled duffel bag.",
         "#PackingTips #CarryOnOnly #MinimalistTravel #WorkTripStyle #TravelCapsule"),

        ("Day 28 (Oct 28)", "Instagram Stories / Interaction", "Pillar 4: Community", 
         "Help Us Choose Our Next Botanical Color Palette!",
         "Community collaboration is at the heart of our design process. For our upcoming Spring Capsule, our dye master has formulated two botanical palettes: Option A: River Indigo & Wild Sage | Option B: Desert Sand & Madder Blossom. Vote in our interactive poll right now—the winning palette goes directly to our dye vats next Monday!",
         "Vote in poll on stories", "Side-by-side swatch cards with natural pigment raw ingredients displayed alongside dipped textile swatches.",
         "#CoCreation #CommunityDesign #VoteNow #SpringPalette #TerraThreads"),

        ("Day 29 (Oct 29)", "Instagram Carousel", "Pillar 1: Transparency", 
         "The Journey of an Indigo Plant: From Seed to Handloom Dip",
         "True indigo (*Indigofera tinctoria*) is living plant alchemy. Swipe through to trace its 90-day sacred metamorphosis: 1) Cultivation in regenerative soil, 2) Leaf harvesting and fermentation, 3) Oxygenation revealing the vibrant blue froth, 4) Successive dip-dyeing by third-generation master dyers. Pure nature, zero synthetic chemicals.",
         "Save this slice of textile history", "Stunning photographic documentary series capturing green indigo leaves transforming into deep cobalt blue vats.",
         "#NaturalIndigo #IndigoDyeing #ArtisanCraft #PlantAlchemy #HeritageTextiles"),

        ("Day 30 (Oct 30)", "Instagram Reel / Milestone Celebration", "Pillar 4: Community", 
         "30 Days of Conscious Living: What We Learned Together",
         "Over the past 30 days, over 50,000 of you engaged with our transparency audits, participated in the #30WearChallenge, and committed to dressing with mindful intention. Sustainability isn't about clinical perfection by a select few; it is about millions of us making mindful, compassionate choices every single morning. Thank you for walking this circular path with us.",
         "Share your favourite takeaway below", "Fast-moving, emotionally inspiring montage of artisan smiles, customer candid photos, and dye studio ripples.",
         "#TerraThreads #SlowFashionMovement #CircularFuture #ThankYou #WearYourValues")
    ]

    # Create Master Table in Document
    tbl = doc.add_table(rows=len(posts) + 1, cols=4)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    
    col_widths = [Inches(1.2), Inches(1.3), Inches(2.2), Inches(1.8)]
    for row in tbl.rows:
        for i, w in enumerate(col_widths):
            row.cells[i].width = w

    headers = ["Day & Date", "Platform & Pillar", "Post Title & Hook", "Creative & CTA"]
    for i, h in enumerate(headers):
        cell = tbl.cell(0, i)
        set_cell_background(cell, "0F1E36")
        set_cell_margins(cell, top=120, bottom=120, left=140, right=140)
        p = cell.paragraphs[0]
        r = p.add_run(h)
        r.font.name = "Calibri"
        r.font.size = Pt(9.5)
        r.font.bold = True
        r.font.color.rgb = RGBColor(56, 189, 248)

    for row_idx, post_item in enumerate(posts):
        p_day, p_plat, p_pill, p_hook, p_cap, p_cta, p_vis, p_tags = post_item
        cell_0 = tbl.cell(row_idx + 1, 0)
        cell_1 = tbl.cell(row_idx + 1, 1)
        cell_2 = tbl.cell(row_idx + 1, 2)
        cell_3 = tbl.cell(row_idx + 1, 3)

        bg = "F8FAFC" if row_idx % 2 == 0 else "FFFFFF"
        for c in (cell_0, cell_1, cell_2, cell_3):
            set_cell_background(c, bg)
            set_cell_margins(c, top=80, bottom=80, left=100, right=100)

        # Col 0: Day
        p = cell_0.paragraphs[0]
        r = p.add_run(p_day)
        r.font.bold = True
        r.font.size = Pt(8.5)
        r.font.color.rgb = RGBColor(15, 23, 42)

        # Col 1: Platform & Pillar
        p = cell_1.paragraphs[0]
        r1 = p.add_run(p_plat + "\n")
        r1.font.bold = True
        r1.font.size = Pt(8)
        r1.font.color.rgb = RGBColor(2, 132, 199)
        r2 = p.add_run(p_pill)
        r2.font.size = Pt(7.5)
        r2.font.color.rgb = RGBColor(100, 116, 139)

        # Col 2: Hook & Caption
        p = cell_2.paragraphs[0]
        r1 = p.add_run(p_hook + "\n")
        r1.font.bold = True
        r1.font.size = Pt(8.5)
        r1.font.color.rgb = RGBColor(15, 23, 42)
        r2 = p.add_run(p_cap[:140] + "...")
        r2.font.size = Pt(7.5)
        r2.font.color.rgb = RGBColor(51, 65, 85)

        # Col 3: Visual & CTA
        p = cell_3.paragraphs[0]
        r1 = p.add_run("CTA: " + p_cta + "\n")
        r1.font.bold = True
        r1.font.size = Pt(8)
        r1.font.color.rgb = RGBColor(217, 119, 6)
        r2 = p.add_run(p_vis[:80] + "...")
        r2.font.size = Pt(7.2)
        r2.font.color.rgb = RGBColor(100, 116, 139)

        # Markdown representation
        log(f"### {p_day} — {p_hook}")
        log(f"- **Platform:** {p_plat} | **Pillar:** {p_pill}")
        log(f"- **Caption:** {p_cap}")
        log(f"- **CTA:** {p_cta} | **Visual Direction:** {p_vis}")
        log(f"- **Hashtags:** `{p_tags}`\n")

    doc.add_paragraph() # Spacer

    # Section 4: Performance Measurement & KPIs
    add_heading_1(doc, "4. Social Media Performance Measurement & KPI Scorecard")
    log("## 4. Social Media Performance Measurement & KPI Scorecard\n")

    p_kpi = ("To evaluate the commercial and community efficacy of this 30-day editorial campaign, performance will be benchmarked "
             "against four quantitative and qualitative metrics: 1) Engagement Rate (Target: > 4.5% on Instagram carousels and reels), "
             "2) Save Rate (Target: > 8% on educational textile infographics and styling lookbooks, indicating high reference utility), "
             "3) Click-Through Rate (CTR to link in bio > 2.2% driving lookbook downloads and ethical transparency audits), and "
             "4) Community Sentiment & Brand Affinity (monitoring organic #30WearChallenge mentions and positive comment ratios).")
    add_body_paragraph(doc, p_kpi)
    log(p_kpi + "\n")

    docx_path = os.path.join(DOCX_DIR, "Week_02_Social_Media_Content_Planning_Sustainable_Fashion.docx")
    doc.save(docx_path)

    md_path = os.path.join(MD_DIR, "Week_02_Social_Media_Content_Planning_Sustainable_Fashion.md")
    with open(md_path, "w", encoding="utf-8") as f:
        f.write("\n".join(md_lines))

    # Calculate word count
    full_text = " ".join([p.text for p in doc.paragraphs])
    for row in tbl.rows:
        for c in row.cells:
            full_text += " " + c.text
    word_count = len(full_text.split())
    print(f"Week 2 Task built successfully: {docx_path} ({word_count} words)")
    return word_count

if __name__ == "__main__":
    build_week_02()
