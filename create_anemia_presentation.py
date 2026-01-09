#!/usr/bin/env python3
"""
Comprehensive PowerPoint Presentation: Systematic Approach to Diagnosing Anemia
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import nsmap
from pptx.oxml import parse_xml

def set_slide_background(slide, r, g, b):
    """Set solid background color for a slide"""
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(r, g, b)

def add_title_slide(prs, title, subtitle):
    """Add a title slide"""
    slide_layout = prs.slide_layouts[6]  # Blank layout
    slide = prs.slides.add_slide(slide_layout)
    set_slide_background(slide, 0, 51, 102)  # Dark blue

    # Title
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(2.5), Inches(9), Inches(1.5))
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    p.text = title
    p.font.size = Pt(44)
    p.font.bold = True
    p.font.color.rgb = RGBColor(255, 255, 255)
    p.alignment = PP_ALIGN.CENTER

    # Subtitle
    subtitle_box = slide.shapes.add_textbox(Inches(0.5), Inches(4.2), Inches(9), Inches(1))
    tf = subtitle_box.text_frame
    p = tf.paragraphs[0]
    p.text = subtitle
    p.font.size = Pt(24)
    p.font.color.rgb = RGBColor(200, 200, 200)
    p.alignment = PP_ALIGN.CENTER

    return slide

def add_section_slide(prs, title):
    """Add a section divider slide"""
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)
    set_slide_background(slide, 139, 0, 0)  # Dark red (blood themed)

    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(2.8), Inches(9), Inches(1.5))
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    p.text = title
    p.font.size = Pt(40)
    p.font.bold = True
    p.font.color.rgb = RGBColor(255, 255, 255)
    p.alignment = PP_ALIGN.CENTER

    return slide

def add_content_slide(prs, title, bullet_points, sub_bullets=None):
    """Add a content slide with title and bullet points"""
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)
    set_slide_background(slide, 255, 255, 255)

    # Add colored header bar
    header = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(10), Inches(1.2))
    header.fill.solid()
    header.fill.fore_color.rgb = RGBColor(0, 51, 102)
    header.line.fill.background()

    # Title
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.8))
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    p.text = title
    p.font.size = Pt(32)
    p.font.bold = True
    p.font.color.rgb = RGBColor(255, 255, 255)

    # Content
    content_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(9), Inches(5.5))
    tf = content_box.text_frame
    tf.word_wrap = True

    for i, point in enumerate(bullet_points):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.text = "• " + point
        p.font.size = Pt(20)
        p.font.color.rgb = RGBColor(0, 0, 0)
        p.space_after = Pt(12)

        # Add sub-bullets if provided
        if sub_bullets and i in sub_bullets:
            for sub in sub_bullets[i]:
                p = tf.add_paragraph()
                p.text = "    ◦ " + sub
                p.font.size = Pt(18)
                p.font.color.rgb = RGBColor(80, 80, 80)
                p.space_after = Pt(6)

    return slide

def add_two_column_slide(prs, title, left_title, left_points, right_title, right_points):
    """Add a two-column content slide"""
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)
    set_slide_background(slide, 255, 255, 255)

    # Header bar
    header = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(10), Inches(1.2))
    header.fill.solid()
    header.fill.fore_color.rgb = RGBColor(0, 51, 102)
    header.line.fill.background()

    # Title
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.8))
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    p.text = title
    p.font.size = Pt(32)
    p.font.bold = True
    p.font.color.rgb = RGBColor(255, 255, 255)

    # Left column title
    left_title_box = slide.shapes.add_textbox(Inches(0.3), Inches(1.4), Inches(4.5), Inches(0.5))
    tf = left_title_box.text_frame
    p = tf.paragraphs[0]
    p.text = left_title
    p.font.size = Pt(22)
    p.font.bold = True
    p.font.color.rgb = RGBColor(139, 0, 0)

    # Left content
    left_box = slide.shapes.add_textbox(Inches(0.3), Inches(1.9), Inches(4.5), Inches(5))
    tf = left_box.text_frame
    tf.word_wrap = True
    for i, point in enumerate(left_points):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.text = "• " + point
        p.font.size = Pt(17)
        p.space_after = Pt(8)

    # Right column title
    right_title_box = slide.shapes.add_textbox(Inches(5.2), Inches(1.4), Inches(4.5), Inches(0.5))
    tf = right_title_box.text_frame
    p = tf.paragraphs[0]
    p.text = right_title
    p.font.size = Pt(22)
    p.font.bold = True
    p.font.color.rgb = RGBColor(139, 0, 0)

    # Right content
    right_box = slide.shapes.add_textbox(Inches(5.2), Inches(1.9), Inches(4.5), Inches(5))
    tf = right_box.text_frame
    tf.word_wrap = True
    for i, point in enumerate(right_points):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.text = "• " + point
        p.font.size = Pt(17)
        p.space_after = Pt(8)

    return slide

def add_table_slide(prs, title, headers, rows):
    """Add a slide with a table"""
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)
    set_slide_background(slide, 255, 255, 255)

    # Header bar
    header_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(10), Inches(1.0))
    header_bar.fill.solid()
    header_bar.fill.fore_color.rgb = RGBColor(0, 51, 102)
    header_bar.line.fill.background()

    # Title
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(9), Inches(0.6))
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    p.text = title
    p.font.size = Pt(28)
    p.font.bold = True
    p.font.color.rgb = RGBColor(255, 255, 255)

    # Table
    num_cols = len(headers)
    num_rows = len(rows) + 1
    table_width = Inches(9)
    table_height = Inches(5.5)

    table = slide.shapes.add_table(num_rows, num_cols, Inches(0.5), Inches(1.2), table_width, table_height).table

    # Set column widths
    col_width = table_width // num_cols
    for i in range(num_cols):
        table.columns[i].width = col_width

    # Header row
    for i, header_text in enumerate(headers):
        cell = table.cell(0, i)
        cell.text = header_text
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(0, 51, 102)
        p = cell.text_frame.paragraphs[0]
        p.font.bold = True
        p.font.size = Pt(14)
        p.font.color.rgb = RGBColor(255, 255, 255)
        p.alignment = PP_ALIGN.CENTER

    # Data rows
    for row_idx, row_data in enumerate(rows):
        for col_idx, cell_text in enumerate(row_data):
            cell = table.cell(row_idx + 1, col_idx)
            cell.text = cell_text
            p = cell.text_frame.paragraphs[0]
            p.font.size = Pt(12)
            p.alignment = PP_ALIGN.LEFT
            # Alternate row colors
            if row_idx % 2 == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(240, 240, 240)

    return slide

def add_algorithm_slide(prs, title, steps):
    """Add a diagnostic algorithm/flowchart slide"""
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)
    set_slide_background(slide, 255, 255, 255)

    # Header bar
    header = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(10), Inches(1.0))
    header.fill.solid()
    header.fill.fore_color.rgb = RGBColor(0, 51, 102)
    header.line.fill.background()

    # Title
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(9), Inches(0.6))
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    p.text = title
    p.font.size = Pt(28)
    p.font.bold = True
    p.font.color.rgb = RGBColor(255, 255, 255)

    # Add flowchart boxes
    y_pos = 1.3
    for i, (step_title, step_content) in enumerate(steps):
        # Box
        box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1), Inches(y_pos), Inches(8), Inches(0.8))
        box.fill.solid()
        if i == 0:
            box.fill.fore_color.rgb = RGBColor(0, 51, 102)  # Dark blue for first
        elif i == len(steps) - 1:
            box.fill.fore_color.rgb = RGBColor(0, 100, 0)  # Green for last
        else:
            box.fill.fore_color.rgb = RGBColor(70, 130, 180)  # Steel blue for middle

        # Text in box
        tf = box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = f"{step_title}: {step_content}"
        p.font.size = Pt(14)
        p.font.bold = True
        p.font.color.rgb = RGBColor(255, 255, 255)
        p.alignment = PP_ALIGN.CENTER
        tf.paragraphs[0].alignment = PP_ALIGN.CENTER

        # Arrow (except after last box)
        if i < len(steps) - 1:
            arrow = slide.shapes.add_shape(MSO_SHAPE.DOWN_ARROW, Inches(4.7), Inches(y_pos + 0.85), Inches(0.6), Inches(0.35))
            arrow.fill.solid()
            arrow.fill.fore_color.rgb = RGBColor(139, 0, 0)
            arrow.line.fill.background()

        y_pos += 1.2

    return slide

def create_presentation():
    """Create the complete anemia diagnosis presentation"""
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)

    # ========== SLIDE 1: Title ==========
    add_title_slide(prs,
        "Systematic Approach to\nDiagnosing Anemia",
        "A Comprehensive Clinical Guide\n2024-2025 Updated Guidelines")

    # ========== SLIDE 2: Learning Objectives ==========
    add_content_slide(prs, "Learning Objectives", [
        "Define anemia and understand WHO diagnostic criteria",
        "Apply a systematic MCV-based approach to anemia diagnosis",
        "Interpret key laboratory findings in anemia workup",
        "Differentiate between common causes of anemia",
        "Recognize indications for specialist referral",
        "Understand the reticulocyte-based kinetic approach",
        "Apply diagnostic algorithms to clinical scenarios"
    ])

    # ========== SLIDE 3: Definition of Anemia ==========
    add_content_slide(prs, "Definition of Anemia", [
        "Reduction in hemoglobin (Hb), hematocrit (Hct), or RBC count below normal",
        "WHO Criteria for Anemia:",
        "Anemia reflects an underlying disease process - not a diagnosis itself",
        "Must consider patient's baseline and clinical context"
    ], sub_bullets={
        1: [
            "Adult males: Hb < 13 g/dL",
            "Adult females (non-pregnant): Hb < 12 g/dL",
            "Pregnant women: Hb < 11 g/dL",
            "Children (6 months - 5 years): Hb < 11 g/dL",
            "Children (5-12 years): Hb < 11.5 g/dL"
        ]
    })

    # ========== SLIDE 4: Epidemiology ==========
    add_content_slide(prs, "Epidemiology & Global Burden", [
        "Affects ~1.8 billion people globally (WHO 2023 data)",
        "Most common hematologic disorder worldwide",
        "Iron deficiency anemia: Most prevalent cause globally",
        "Higher prevalence in developing countries",
        "Risk groups: Women of reproductive age, children, elderly",
        "Associated with increased morbidity and mortality",
        "Significant economic burden on healthcare systems"
    ])

    # ========== SECTION: Clinical Assessment ==========
    add_section_slide(prs, "SECTION 1\nClinical Assessment")

    # ========== SLIDE 5: Clinical Presentation ==========
    add_two_column_slide(prs, "Clinical Presentation of Anemia",
        "General Symptoms",
        [
            "Fatigue and weakness",
            "Dyspnea on exertion",
            "Palpitations",
            "Dizziness/lightheadedness",
            "Headache",
            "Decreased exercise tolerance",
            "Chest pain (severe anemia)",
            "Pallor (conjunctivae, palms, nail beds)"
        ],
        "Specific Signs by Etiology",
        [
            "Koilonychia (iron deficiency)",
            "Glossitis, angular cheilitis",
            "Jaundice (hemolysis)",
            "Splenomegaly",
            "Petechiae (marrow failure)",
            "Neurological signs (B12 deficiency)",
            "Leg ulcers (sickle cell)",
            "Bone deformities (thalassemia major)"
        ]
    )

    # ========== SLIDE 6: History Taking ==========
    add_content_slide(prs, "Essential History Taking", [
        "Duration and onset of symptoms",
        "Dietary history (vegetarian/vegan, pica, nutrition)",
        "Bleeding history (menstrual, GI, urinary)",
        "Medical conditions (chronic disease, malignancy, autoimmune)",
        "Medications (NSAIDs, anticoagulants, chemotherapy)",
        "Family history (hemoglobinopathies, hereditary anemias)",
        "Ethnicity (G6PD, thalassemia, sickle cell)",
        "Alcohol use, drug history",
        "Surgical history (gastric bypass, resections)"
    ])

    # ========== SECTION: Laboratory Evaluation ==========
    add_section_slide(prs, "SECTION 2\nLaboratory Evaluation")

    # ========== SLIDE 7: Initial Lab Workup ==========
    add_content_slide(prs, "Initial Laboratory Workup", [
        "Complete Blood Count (CBC) with indices",
        "Reticulocyte count (absolute and percentage)",
        "Peripheral blood smear examination",
        "Basic metabolic panel",
        "Iron studies (serum iron, TIBC, ferritin, transferrin saturation)",
        "Vitamin B12 and folate levels",
        "Lactate dehydrogenase (LDH)",
        "Bilirubin (total and indirect)",
        "Haptoglobin (if hemolysis suspected)"
    ])

    # ========== SLIDE 8: CBC Indices ==========
    add_table_slide(prs, "Key CBC Parameters & Reference Ranges",
        ["Parameter", "Normal Range", "Clinical Significance"],
        [
            ["Hemoglobin (Hb)", "M: 13-17 g/dL\nF: 12-16 g/dL", "Primary marker of anemia severity"],
            ["Hematocrit (Hct)", "M: 40-54%\nF: 36-48%", "Reflects RBC volume fraction"],
            ["MCV", "80-100 fL", "Classifies anemia morphologically"],
            ["MCH", "27-33 pg", "Hb content per RBC"],
            ["MCHC", "32-36 g/dL", "Hb concentration per RBC"],
            ["RDW", "11.5-14.5%", "Variation in RBC size (anisocytosis)"],
            ["Reticulocytes", "0.5-2.5%", "Bone marrow response indicator"]
        ]
    )

    # ========== SECTION: MCV-Based Classification ==========
    add_section_slide(prs, "SECTION 3\nMCV-Based Classification")

    # ========== SLIDE 9: MCV Classification Overview ==========
    add_content_slide(prs, "MCV-Based Classification of Anemia", [
        "MICROCYTIC (MCV < 80 fL):",
        "NORMOCYTIC (MCV 80-100 fL):",
        "MACROCYTIC (MCV > 100 fL):"
    ], sub_bullets={
        0: ["Iron deficiency", "Thalassemia", "Anemia of chronic disease", "Sideroblastic anemia", "Lead poisoning"],
        1: ["Acute blood loss", "Anemia of chronic disease", "Chronic kidney disease", "Hemolysis", "Bone marrow failure", "Mixed deficiencies"],
        2: ["Vitamin B12 deficiency", "Folate deficiency", "Myelodysplastic syndrome", "Liver disease", "Hypothyroidism", "Medications", "Reticulocytosis"]
    })

    # ========== SLIDE 10: Microcytic Anemia Approach ==========
    add_table_slide(prs, "Differentiating Microcytic Anemias",
        ["Parameter", "Iron Deficiency", "Thalassemia Trait", "ACD", "Sideroblastic"],
        [
            ["Serum Iron", "↓↓", "Normal", "↓", "↑"],
            ["TIBC", "↑↑", "Normal", "↓", "Normal/↓"],
            ["Ferritin", "↓↓ (<30)", "Normal", "Normal/↑", "↑↑"],
            ["Transferrin Sat", "< 16%", "Normal", "Normal/↓", "↑"],
            ["RDW", "↑↑", "Normal", "Normal", "↑"],
            ["RBC Count", "↓", "↑ or Normal", "↓", "↓"],
            ["Mentzer Index", "> 13", "< 13", "Variable", "Variable"],
            ["Special Tests", "Low ferritin\nsoluble TfR ↑", "Hb electro-\nphoresis", "CRP/ESR ↑\nHepcidin ↑", "Ring sidero-\nblasts on BM"]
        ]
    )

    # ========== SLIDE 11: Iron Deficiency Workup ==========
    add_content_slide(prs, "Iron Deficiency Anemia - Workup", [
        "Confirm iron deficiency: Ferritin < 30 ng/mL (< 100 in inflammation)",
        "Identify the cause - Iron deficiency is a symptom!",
        "Common causes to investigate:",
        "Further workup based on clinical suspicion:",
        "Consider GI evaluation in all men and postmenopausal women"
    ], sub_bullets={
        2: [
            "GI blood loss (most common in men/postmenopausal women)",
            "Menstrual blood loss (most common in premenopausal women)",
            "Malabsorption (celiac, H. pylori, gastric surgery)",
            "Dietary insufficiency",
            "Increased demands (pregnancy, growth)"
        ],
        3: [
            "Upper and lower GI endoscopy",
            "Celiac serology (TTG-IgA)",
            "H. pylori testing",
            "Urine analysis for hematuria"
        ]
    })

    # ========== SLIDE 12: Normocytic Anemia ==========
    add_content_slide(prs, "Normocytic Anemia - Diagnostic Approach", [
        "Check reticulocyte count first:",
        "Key considerations:",
        "Essential workup includes:",
        "Consider bone marrow biopsy if unexplained"
    ], sub_bullets={
        0: [
            "High (>2-3%): Hemolysis or acute blood loss",
            "Low/Normal (<2%): Underproduction (marrow failure, chronic disease)"
        ],
        1: [
            "Can be early iron/B12/folate deficiency",
            "May be combined deficiencies masking MCV changes",
            "Anemia of chronic disease can be normocytic or microcytic"
        ],
        2: [
            "Renal function (CKD → low EPO)",
            "Thyroid function",
            "Hemolysis workup (LDH, haptoglobin, bilirubin, DAT)",
            "Peripheral smear review"
        ]
    })

    # ========== SLIDE 13: Macrocytic Anemia ==========
    add_two_column_slide(prs, "Macrocytic Anemia Classification",
        "Megaloblastic",
        [
            "Vitamin B12 deficiency",
            "Folate deficiency",
            "Drug-induced (methotrexate,\n  hydroxyurea, azathioprine)",
            "Features: Hypersegmented\n  neutrophils, oval macrocytes",
            "Ineffective erythropoiesis",
            "Check B12, folate, MMA,\n  homocysteine levels"
        ],
        "Non-Megaloblastic",
        [
            "Liver disease/Alcohol",
            "Hypothyroidism",
            "Myelodysplastic syndrome",
            "Reticulocytosis",
            "Drug-induced (non-antimetabolite)",
            "Round macrocytes",
            "Target cells in liver disease",
            "May need bone marrow biopsy"
        ]
    )

    # ========== SLIDE 14: B12 and Folate Deficiency ==========
    add_table_slide(prs, "B12 vs Folate Deficiency",
        ["Feature", "B12 Deficiency", "Folate Deficiency"],
        [
            ["Common Causes", "Pernicious anemia, malabsorption,\nvegan diet, medications", "Poor diet, alcoholism, pregnancy,\nmalabsorption, medications"],
            ["Neurological Sx", "Present (SACD*, peripheral\nneuropathy, cognitive changes)", "Absent (unless severe)"],
            ["Serum Level", "↓ (< 200 pg/mL)", "↓ (< 3 ng/mL)"],
            ["Methylmalonic Acid", "↑ (specific for B12)", "Normal"],
            ["Homocysteine", "↑", "↑"],
            ["Additional Tests", "Anti-IF antibodies, anti-parietal\ncell Ab, gastrin level", "RBC folate (more reliable\nthan serum)"],
            ["Treatment", "B12 replacement (IM or high-\ndose oral)", "Folate supplementation\n(rule out B12 def first!)"]
        ]
    )

    # ========== SECTION: Hemolytic Anemia ==========
    add_section_slide(prs, "SECTION 4\nHemolytic Anemia")

    # ========== SLIDE 15: Hemolysis Overview ==========
    add_content_slide(prs, "Recognizing Hemolytic Anemia", [
        "Hallmark: Elevated reticulocyte count + anemia",
        "Laboratory findings suggesting hemolysis:",
        "Peripheral smear findings:",
        "Classification: Intrinsic (RBC defect) vs Extrinsic (external cause)"
    ], sub_bullets={
        1: [
            "↑ LDH (released from lysed RBCs)",
            "↓ Haptoglobin (binds free Hb → cleared)",
            "↑ Indirect bilirubin (unconjugated)",
            "↑ Reticulocyte count (compensatory)",
            "Hemoglobinuria/hemosiderinuria (intravascular)"
        ],
        2: [
            "Spherocytes, schistocytes, bite cells",
            "Sickle cells, target cells",
            "RBC agglutination, polychromasia"
        ]
    })

    # ========== SLIDE 16: Hemolysis Classification ==========
    add_two_column_slide(prs, "Classification of Hemolytic Anemias",
        "Intrinsic (RBC Defects)",
        [
            "Membrane disorders:",
            "  - Hereditary spherocytosis",
            "  - Hereditary elliptocytosis",
            "Enzyme deficiencies:",
            "  - G6PD deficiency",
            "  - Pyruvate kinase deficiency",
            "Hemoglobinopathies:",
            "  - Sickle cell disease",
            "  - Thalassemia (major)",
            "  - Unstable hemoglobins"
        ],
        "Extrinsic (External Factors)",
        [
            "Immune-mediated:",
            "  - Autoimmune (warm/cold)",
            "  - Drug-induced",
            "  - Transfusion reactions",
            "Microangiopathic:",
            "  - TTP, HUS, DIC, HELLP",
            "  - Mechanical heart valves",
            "Infections: Malaria, Babesia",
            "Toxins: Snake venom, copper",
            "Hypersplenism"
        ]
    )

    # ========== SLIDE 17: Hemolysis Workup ==========
    add_content_slide(prs, "Hemolytic Anemia Workup", [
        "Step 1: Confirm hemolysis (↑retics, ↓haptoglobin, ↑LDH, ↑indirect bili)",
        "Step 2: Direct Antiglobulin Test (Coombs)",
        "Step 3: Peripheral blood smear - Critical for diagnosis!",
        "Step 4: Additional testing based on smear/clinical context",
        "Step 5: Consider intravascular vs extravascular hemolysis"
    ], sub_bullets={
        1: [
            "DAT positive → Immune hemolysis (AIHA, drug-induced, transfusion)",
            "DAT negative → Non-immune causes"
        ],
        3: [
            "Spherocytes: HS, AIHA",
            "Schistocytes: MAHA (TTP/HUS/DIC)",
            "Bite/blister cells: G6PD deficiency",
            "Sickle cells: SCD",
            "Target cells: Thalassemia, liver disease"
        ]
    })

    # ========== SECTION: Kinetic Approach ==========
    add_section_slide(prs, "SECTION 5\nReticulocyte-Based Approach")

    # ========== SLIDE 18: Reticulocyte Index ==========
    add_content_slide(prs, "Reticulocyte Count & Reticulocyte Index", [
        "Reticulocyte count reflects bone marrow response",
        "Must correct for degree of anemia:",
        "Interpretation of Reticulocyte Production Index (RPI):",
        "Absolute Reticulocyte Count (ARC) > 100,000/μL suggests adequate response",
        "Combines with MCV for comprehensive classification"
    ], sub_bullets={
        1: [
            "Corrected retic count = Retic % × (Patient Hct/Normal Hct)",
            "RPI = Corrected retic count / Maturation factor",
            "Maturation factor: 1 (Hct 45%), 1.5 (Hct 35%), 2 (Hct 25%), 2.5 (Hct 15%)"
        ],
        2: [
            "RPI ≥ 2-3: Appropriate response (hemolysis, blood loss)",
            "RPI < 2: Hypoproliferative (production problem)"
        ]
    })

    # ========== SLIDE 19: Kinetic Algorithm ==========
    add_algorithm_slide(prs, "Kinetic (Reticulocyte-Based) Algorithm", [
        ("Step 1", "Confirm anemia with CBC"),
        ("Step 2", "Calculate RPI or absolute reticulocyte count"),
        ("Step 3a", "RPI ≥ 2: Blood loss or Hemolysis → Hemolysis workup"),
        ("Step 3b", "RPI < 2: Underproduction → Use MCV classification"),
        ("Step 4", "Integrate clinical findings with lab results"),
        ("Step 5", "Targeted additional testing based on differential")
    ])

    # ========== SECTION: Special Situations ==========
    add_section_slide(prs, "SECTION 6\nSpecial Situations")

    # ========== SLIDE 20: Anemia of Chronic Disease ==========
    add_content_slide(prs, "Anemia of Chronic Disease (ACD)", [
        "Second most common cause of anemia worldwide",
        "Pathophysiology: Hepcidin-mediated iron sequestration",
        "Associated conditions:",
        "Laboratory profile:",
        "May coexist with iron deficiency - challenge to diagnose",
        "Soluble transferrin receptor (sTfR) helps differentiate from IDA"
    ], sub_bullets={
        2: [
            "Chronic infections (HIV, TB, osteomyelitis)",
            "Autoimmune/inflammatory disorders (RA, SLE, IBD)",
            "Malignancy",
            "Chronic kidney disease"
        ],
        3: [
            "Usually normocytic (can be microcytic)",
            "Low serum iron, low TIBC",
            "Normal or elevated ferritin (acute phase reactant)",
            "Low-normal transferrin saturation"
        ]
    })

    # ========== SLIDE 21: Anemia in CKD ==========
    add_content_slide(prs, "Anemia of Chronic Kidney Disease", [
        "Primary mechanism: Decreased erythropoietin (EPO) production",
        "Usually develops when eGFR < 60 mL/min/1.73m²",
        "Typically normocytic, normochromic",
        "Contributing factors:",
        "Workup: Rule out other causes before attributing to CKD",
        "Management: Iron supplementation, ESA therapy if indicated",
        "Target Hb: Generally 10-11.5 g/dL (avoid > 13 g/dL)"
    ], sub_bullets={
        3: [
            "Iron deficiency (uremic blood loss, dialysis losses)",
            "Inflammation/chronic disease",
            "Folate/B12 deficiency",
            "Hyperparathyroidism (marrow fibrosis)",
            "Shortened RBC survival"
        ]
    })

    # ========== SLIDE 22: Pregnancy ==========
    add_content_slide(prs, "Anemia in Pregnancy", [
        "Physiological anemia: Plasma volume expansion > RBC increase",
        "WHO criteria: Hb < 11 g/dL (1st/3rd trimester), < 10.5 g/dL (2nd)",
        "Most common causes: Iron deficiency (most common), folate deficiency",
        "Screening recommendations:",
        "Special considerations:",
        "Treatment: Oral iron first-line, IV iron if intolerant/severe"
    ], sub_bullets={
        3: [
            "CBC at first prenatal visit and 24-28 weeks",
            "Iron studies if anemia present",
            "All pregnant women should receive iron supplementation"
        ],
        4: [
            "B12 deficiency can cause neural tube defects",
            "Sickle cell screening in appropriate populations",
            "HELLP syndrome: Microangiopathic hemolysis in preeclampsia"
        ]
    })

    # ========== SLIDE 23: Elderly ==========
    add_content_slide(prs, "Anemia in the Elderly", [
        "Prevalence: ~10% of community-dwelling adults >65, ~50% in nursing homes",
        "Often multifactorial - requires comprehensive evaluation",
        "Common causes:",
        "Unexplained anemia of elderly (UAE):",
        "Associated with increased mortality, falls, cognitive decline",
        "Lower threshold for bone marrow evaluation"
    ], sub_bullets={
        2: [
            "Anemia of chronic disease/inflammation",
            "Iron deficiency (often GI blood loss - malignancy!)",
            "CKD-associated anemia",
            "Nutritional deficiencies (B12, folate)",
            "Myelodysplastic syndrome (more common with age)"
        ],
        3: [
            "~1/3 of elderly anemia cases remain unexplained",
            "May involve age-related decline in hematopoiesis",
            "Consider clonal hematopoiesis of indeterminate potential (CHIP)"
        ]
    })

    # ========== SECTION: Diagnostic Algorithms ==========
    add_section_slide(prs, "SECTION 7\nDiagnostic Algorithms")

    # ========== SLIDE 24: Master Algorithm ==========
    add_algorithm_slide(prs, "Master Diagnostic Algorithm", [
        ("Confirm", "Hb below threshold for age/sex"),
        ("History", "Bleeding, diet, meds, comorbidities, family hx"),
        ("CBC", "Evaluate MCV, RDW, other cell lines"),
        ("Reticulocytes", "Assess bone marrow response"),
        ("Initial Labs", "Iron studies, B12, folate, LFTs, renal function"),
        ("Targeted Tests", "Based on initial findings"),
        ("Diagnosis", "Identify etiology and treat underlying cause")
    ])

    # ========== SLIDE 25: When to Refer ==========
    add_content_slide(prs, "Indications for Hematology Referral", [
        "Unexplained anemia despite initial workup",
        "Suspected bone marrow failure or MDS",
        "Hemolytic anemia (especially DAT-positive)",
        "Suspected hemoglobinopathy requiring confirmation",
        "Multiple cytopenias (bicytopenia, pancytopenia)",
        "Need for bone marrow biopsy",
        "Refractory anemia not responding to treatment",
        "Transfusion-dependent anemia",
        "Suspected clonal disorder"
    ])

    # ========== SLIDE 26: Bone Marrow Indications ==========
    add_content_slide(prs, "Indications for Bone Marrow Examination", [
        "Unexplained cytopenias (bi- or pancytopenia)",
        "Suspected hematologic malignancy",
        "Evaluation for myelodysplastic syndrome",
        "Unexplained anemia after non-invasive workup",
        "Staging of lymphoma or other malignancies",
        "Evaluation for iron stores (gold standard)",
        "Suspected infiltrative disease",
        "Unexplained leukoerythroblastic blood smear",
        "Fever of unknown origin with cytopenias"
    ])

    # ========== SECTION: Summary ==========
    add_section_slide(prs, "SECTION 8\nSummary & Key Points")

    # ========== SLIDE 27: Key Takeaways ==========
    add_content_slide(prs, "Key Takeaways", [
        "Anemia is a sign of underlying disease - always find the cause",
        "Use a systematic approach: History → CBC → MCV classification",
        "Reticulocyte count differentiates production vs destruction problems",
        "Iron studies are essential for microcytic anemia workup",
        "Always consider combined deficiencies in normocytic anemia",
        "Peripheral blood smear is invaluable and often underutilized",
        "Don't forget to rule out GI malignancy in iron deficiency",
        "Know when to refer to hematology",
        "Treat the underlying cause, not just the anemia"
    ])

    # ========== SLIDE 28: Quick Reference ==========
    add_table_slide(prs, "Quick Reference: First-Line Tests by MCV",
        ["MCV Category", "First-Line Tests", "Consider If Unrevealing"],
        [
            ["Microcytic\n(< 80 fL)", "Iron studies (ferritin, Fe, TIBC,\nTSAT), Hb electrophoresis", "Lead level, soluble TfR,\nbone marrow biopsy"],
            ["Normocytic\n(80-100 fL)", "Reticulocyte count, iron studies,\nB12/folate, creatinine, TSH", "Hemolysis workup, DAT,\nBM biopsy, EPO level"],
            ["Macrocytic\n(> 100 fL)", "B12, folate, TSH, LFTs,\nreticulocyte count, smear", "MMA, homocysteine,\nBM biopsy (MDS workup)"]
        ]
    )

    # ========== SLIDE 29: References ==========
    add_content_slide(prs, "References & Further Reading", [
        "WHO Haemoglobin Concentrations for Diagnosis of Anaemia (2024 update)",
        "American Society of Hematology Clinical Guidelines",
        "UpToDate: Approach to the adult with anemia",
        "Harrison's Principles of Internal Medicine, 21st Ed (2022)",
        "Blood Journal - ASH Education Program",
        "British Journal of Haematology Guidelines",
        "KDIGO Clinical Practice Guidelines for Anemia in CKD",
        "American College of Gastroenterology Guidelines on GI Evaluation"
    ])

    # ========== SLIDE 30: Thank You ==========
    add_title_slide(prs,
        "Thank You",
        "Questions & Discussion")

    # Save presentation
    prs.save('Systematic_Approach_to_Diagnosing_Anemia.pptx')
    print("Presentation created successfully: Systematic_Approach_to_Diagnosing_Anemia.pptx")

if __name__ == "__main__":
    create_presentation()
