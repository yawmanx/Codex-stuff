#!/usr/bin/env python3
"""
PowerPoint-Präsentation: Komplikationen des Magenbypasses
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

def add_title_slide(prs, title, subtitle=""):
    """Add a title slide"""
    slide_layout = prs.slide_layouts[6]  # Blank layout
    slide = prs.slides.add_slide(slide_layout)

    # Add background shape
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(0, 82, 147)  # Medical blue
    shape.line.fill.background()

    # Title
    txBox = slide.shapes.add_textbox(Inches(0.5), Inches(2.5), Inches(9), Inches(1.5))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = title
    p.font.size = Pt(44)
    p.font.bold = True
    p.font.color.rgb = RGBColor(255, 255, 255)
    p.alignment = PP_ALIGN.CENTER

    # Subtitle
    if subtitle:
        txBox2 = slide.shapes.add_textbox(Inches(0.5), Inches(4), Inches(9), Inches(1))
        tf2 = txBox2.text_frame
        p2 = tf2.paragraphs[0]
        p2.text = subtitle
        p2.font.size = Pt(24)
        p2.font.color.rgb = RGBColor(200, 220, 255)
        p2.alignment = PP_ALIGN.CENTER

    return slide

def add_content_slide(prs, title, bullet_points, highlight_color=None):
    """Add a content slide with bullet points"""
    slide_layout = prs.slide_layouts[6]  # Blank layout
    slide = prs.slides.add_slide(slide_layout)

    # Header bar
    header = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, Inches(1.2))
    header.fill.solid()
    header.fill.fore_color.rgb = RGBColor(0, 82, 147)
    header.line.fill.background()

    # Title
    txBox = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.8))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = title
    p.font.size = Pt(32)
    p.font.bold = True
    p.font.color.rgb = RGBColor(255, 255, 255)

    # Bullet points
    content_box = slide.shapes.add_textbox(Inches(0.7), Inches(1.5), Inches(8.6), Inches(5.5))
    tf = content_box.text_frame
    tf.word_wrap = True

    for i, point in enumerate(bullet_points):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()

        p.text = f"• {point}"
        p.font.size = Pt(20)
        p.font.color.rgb = RGBColor(51, 51, 51)
        p.space_after = Pt(12)
        p.level = 0

    return slide

def add_section_slide(prs, title, number=""):
    """Add a section divider slide"""
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)

    # Background
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(220, 53, 69)  # Red accent
    shape.line.fill.background()

    # Number
    if number:
        txBox = slide.shapes.add_textbox(Inches(0.5), Inches(2), Inches(9), Inches(1))
        tf = txBox.text_frame
        p = tf.paragraphs[0]
        p.text = number
        p.font.size = Pt(72)
        p.font.bold = True
        p.font.color.rgb = RGBColor(255, 255, 255)
        p.alignment = PP_ALIGN.CENTER

    # Title
    txBox2 = slide.shapes.add_textbox(Inches(0.5), Inches(3.2), Inches(9), Inches(1.5))
    tf2 = txBox2.text_frame
    p2 = tf2.paragraphs[0]
    p2.text = title
    p2.font.size = Pt(40)
    p2.font.bold = True
    p2.font.color.rgb = RGBColor(255, 255, 255)
    p2.alignment = PP_ALIGN.CENTER

    return slide

def add_two_column_slide(prs, title, left_title, left_points, right_title, right_points):
    """Add a two-column content slide"""
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)

    # Header bar
    header = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, Inches(1.2))
    header.fill.solid()
    header.fill.fore_color.rgb = RGBColor(0, 82, 147)
    header.line.fill.background()

    # Title
    txBox = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.8))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = title
    p.font.size = Pt(32)
    p.font.bold = True
    p.font.color.rgb = RGBColor(255, 255, 255)

    # Left column title
    left_title_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(4.3), Inches(0.5))
    tf = left_title_box.text_frame
    p = tf.paragraphs[0]
    p.text = left_title
    p.font.size = Pt(22)
    p.font.bold = True
    p.font.color.rgb = RGBColor(0, 82, 147)

    # Left column content
    left_box = slide.shapes.add_textbox(Inches(0.5), Inches(2.1), Inches(4.3), Inches(4.5))
    tf = left_box.text_frame
    tf.word_wrap = True
    for i, point in enumerate(left_points):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.text = f"• {point}"
        p.font.size = Pt(16)
        p.font.color.rgb = RGBColor(51, 51, 51)
        p.space_after = Pt(8)

    # Right column title
    right_title_box = slide.shapes.add_textbox(Inches(5.2), Inches(1.5), Inches(4.3), Inches(0.5))
    tf = right_title_box.text_frame
    p = tf.paragraphs[0]
    p.text = right_title
    p.font.size = Pt(22)
    p.font.bold = True
    p.font.color.rgb = RGBColor(220, 53, 69)

    # Right column content
    right_box = slide.shapes.add_textbox(Inches(5.2), Inches(2.1), Inches(4.3), Inches(4.5))
    tf = right_box.text_frame
    tf.word_wrap = True
    for i, point in enumerate(right_points):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.text = f"• {point}"
        p.font.size = Pt(16)
        p.font.color.rgb = RGBColor(51, 51, 51)
        p.space_after = Pt(8)

    return slide

def create_presentation():
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)

    # Slide 1: Title
    add_title_slide(prs,
        "Komplikationen des Magenbypasses",
        "Ein umfassender Überblick für medizinisches Fachpersonal")

    # Slide 2: Agenda
    add_content_slide(prs, "Übersicht", [
        "Einführung in den Magenbypass",
        "Klassifikation der Komplikationen",
        "Frühe postoperative Komplikationen",
        "Späte Komplikationen",
        "Ernährungsbedingte Mangelzustände",
        "Psychologische Aspekte",
        "Prävention und Management",
        "Zusammenfassung"
    ])

    # Slide 3: Introduction
    add_content_slide(prs, "Einführung: Der Magenbypass", [
        "Roux-en-Y-Magenbypass (RYGB) ist der Goldstandard",
        "Kombiniert restriktive und malabsorptive Mechanismen",
        "Effektiver Gewichtsverlust: 60-80% des Übergewichts",
        "Verbesserung von Komorbiditäten (Diabetes, Hypertonie)",
        "Komplikationsrate: 10-20% insgesamt",
        "Mortalitätsrate: < 0,5% in erfahrenen Zentren"
    ])

    # Slide 4: Classification
    add_two_column_slide(prs, "Klassifikation der Komplikationen",
        "Frühe Komplikationen (< 30 Tage)",
        ["Anastomoseninsuffizienz", "Blutung", "Lungenembolie",
         "Wundinfektion", "Darmobstruktion"],
        "Späte Komplikationen (> 30 Tage)",
        ["Dumping-Syndrom", "Nährstoffmängel", "Stenosen",
         "Innere Hernien", "Ulzera"])

    # Slide 5: Section - Early Complications
    add_section_slide(prs, "Frühe Postoperative Komplikationen", "01")

    # Slide 6: Anastomotic Leak
    add_content_slide(prs, "Anastomoseninsuffizienz", [
        "Inzidenz: 1-5% der Fälle",
        "Häufigste Lokalisation: Gastrojejunostomie",
        "Symptome: Tachykardie, Fieber, Bauchschmerzen, Sepsis",
        "Diagnose: CT mit oralem Kontrastmittel",
        "Therapie: Drainage, Antibiotika, ggf. Re-Operation",
        "Mortalitätsrisiko ohne Behandlung: sehr hoch"
    ])

    # Slide 7: Bleeding
    add_content_slide(prs, "Postoperative Blutung", [
        "Inzidenz: 1-4% der Patienten",
        "Intraluminal (GI-Blutung) vs. Extraluminal (intraabdominal)",
        "Risikofaktoren: Antikoagulation, technische Faktoren",
        "Symptome: Hämatemesis, Meläna, Hämodynamische Instabilität",
        "Diagnose: Endoskopie, CT-Angiographie",
        "Therapie: Endoskopische Blutstillung, Transfusion, ggf. Re-OP"
    ])

    # Slide 8: Thromboembolism
    add_content_slide(prs, "Thromboembolische Komplikationen", [
        "Tiefe Venenthrombose (TVT): 0,3-1,2%",
        "Lungenembolie (LE): 0,2-1%",
        "Haupttodesursache nach bariatrischer Chirurgie",
        "Risikofaktoren: Adipositas, Immobilisation, lange OP-Zeit",
        "Prävention: Frühmobilisation, Kompressionsstrümpfe, Heparin",
        "Therapie: Antikoagulation, ggf. Lyse oder Thrombektomie"
    ])

    # Slide 9: Infection
    add_content_slide(prs, "Infektiöse Komplikationen", [
        "Wundinfektion: 2-8% (seltener bei laparoskopischem Zugang)",
        "Intraabdomineller Abszess: 1-2%",
        "Pneumonie: 0,5-2%",
        "Risikofaktoren: Diabetes, Immunsuppression, lange OP-Dauer",
        "Prävention: Perioperative Antibiotikaprophylaxe",
        "Therapie: Antibiotika, Drainage, Wundpflege"
    ])

    # Slide 10: Section - Late Complications
    add_section_slide(prs, "Späte Komplikationen", "02")

    # Slide 11: Dumping Syndrome
    add_two_column_slide(prs, "Dumping-Syndrom",
        "Frühdumping (15-30 min)",
        ["Übelkeit, Erbrechen, Krämpfe",
         "Diarrhoe, Blähungen",
         "Schwitzen, Tachykardie",
         "Ursache: Schnelle Magenentleerung"],
        "Spätdumping (1-3 Stunden)",
        ["Hypoglykämie",
         "Schwäche, Zittern",
         "Schweißausbrüche",
         "Ursache: Überschießende Insulinsekretion"])

    # Slide 12: Dumping Management
    add_content_slide(prs, "Management des Dumping-Syndroms", [
        "Prävalenz: 20-50% der Bypass-Patienten",
        "Ernährungsmodifikation ist Therapie der ersten Wahl",
        "Kleine, häufige Mahlzeiten (5-6 pro Tag)",
        "Reduktion von Einfachzuckern und raffinierten Kohlenhydraten",
        "Trennung von Essen und Trinken (30 min Abstand)",
        "Medikamentös: Acarbose, Octreotid bei therapierefraktären Fällen"
    ])

    # Slide 13: Section - Nutritional Deficiencies
    add_section_slide(prs, "Ernährungsbedingte Mangelzustände", "03")

    # Slide 14: Vitamin Deficiencies
    add_content_slide(prs, "Vitaminmangelzustände", [
        "Vitamin B12: 30-70% (fehlender Intrinsic-Faktor)",
        "Vitamin D: 50-80% (Malabsorption, wenig Sonnenlicht)",
        "Folsäure: 15-35%",
        "Vitamin B1 (Thiamin): 10-20% (Wernicke-Enzephalopathie!)",
        "Vitamin A: 10-15%",
        "Lebenslange Supplementierung erforderlich"
    ])

    # Slide 15: Mineral Deficiencies
    add_content_slide(prs, "Mineralstoffmängel", [
        "Eisenmangel: 20-50% (häufigste Anämieursache)",
        "Kalziummangel: 10-25% (Osteoporose-Risiko)",
        "Zinkmangel: 10-40% (Haarausfall, Wundheilungsstörungen)",
        "Magnesium: 5-15%",
        "Kupfermangel: selten, aber schwere neurologische Folgen",
        "Regelmäßige Laborkontrollen sind essentiell"
    ])

    # Slide 16: Stenosis
    add_content_slide(prs, "Anastomosenstenose", [
        "Inzidenz: 3-12% an der Gastrojejunostomie",
        "Typische Präsentation: 4-8 Wochen postoperativ",
        "Symptome: Dysphagie, Übelkeit, Erbrechen nach Mahlzeiten",
        "Ursachen: Ischämie, Narbenbildung, technische Faktoren",
        "Diagnose: Endoskopie mit Strikturnachweis",
        "Therapie: Endoskopische Ballondilatation (oft mehrfach nötig)"
    ])

    # Slide 17: Internal Hernia
    add_content_slide(prs, "Innere Hernien", [
        "Inzidenz: 2-5% (häufiger nach laparoskopischem Bypass)",
        "Entstehung durch Mesenteriallücken (Petersen-Hernie)",
        "Symptome: Kolikartige Bauchschmerzen, oft postprandial",
        "Gefahr: Darminkarzeration mit Ischämie und Nekrose",
        "Diagnose: CT-Abdomen (Wirbelzeichen des Mesenteriums)",
        "Therapie: Operative Revision (oft laparoskopisch)"
    ])

    # Slide 18: Marginal Ulcer
    add_content_slide(prs, "Marginalulzera (Anastomosenulzera)", [
        "Inzidenz: 1-16% der Patienten",
        "Lokalisation: Jejunale Seite der Gastrojejunostomie",
        "Risikofaktoren: Rauchen, NSAR, H. pylori, große Pouch",
        "Symptome: Epigastrische Schmerzen, Übelkeit, GI-Blutung",
        "Diagnose: Ösophagogastroduodenoskopie",
        "Therapie: PPI-Hochdosis, Raucherentwöhnung, H. pylori-Eradikation"
    ])

    # Slide 19: Gallstones
    add_content_slide(prs, "Cholelithiasis nach Magenbypass", [
        "Inzidenz: 30-40% entwickeln Gallensteine nach Bypass",
        "Ursache: Schneller Gewichtsverlust → Cholesterinübersättigung",
        "Symptome: Rechtsseitige Oberbauchschmerzen, Koliken",
        "Prävention: Ursodeoxycholsäure in den ersten 6 Monaten",
        "Therapie: Laparoskopische Cholezystektomie bei Symptomen",
        "Einige Zentren führen prophylaktische Cholezystektomie durch"
    ])

    # Slide 20: Psychological Complications
    add_content_slide(prs, "Psychologische Komplikationen", [
        "Depressionen: Erhöhtes Risiko in den ersten Jahren",
        "Suchtverschiebung: Alkohol, Medikamente, Kaufsucht",
        "Essstörungen: Transfer Addiction Syndrome",
        "Beziehungsprobleme nach drastischer Gewichtsabnahme",
        "Suizidalität: Erhöhtes Risiko (besonders Jahre 2-5)",
        "Lebenslange psychologische Betreuung empfohlen"
    ])

    # Slide 21: Prevention
    add_content_slide(prs, "Prävention und Langzeitmanagement", [
        "Präoperative Patientenselektion und -edukation",
        "Strukturiertes Nachsorgeprogramm (lebenslang)",
        "Regelmäßige Laborkontrollen: alle 3-6 Monate initial",
        "Multivitamin- und Mineralstoffsupplementierung",
        "Ernährungsberatung und Verhaltenstherapie",
        "Interdisziplinäres Team: Chirurg, Internist, Ernährungsberater, Psychologe"
    ])

    # Slide 22: Summary
    add_title_slide(prs,
        "Zusammenfassung",
        "")

    # Slide 23: Key Takeaways
    add_content_slide(prs, "Kernbotschaften", [
        "Magenbypass ist effektiv, aber mit Risiken verbunden",
        "Frühe Komplikationen erfordern sofortiges Handeln",
        "Späte Komplikationen erfordern Vigilanz und Nachsorge",
        "Nährstoffmängel sind häufig und behandelbar",
        "Lebenslange Nachsorge ist unerlässlich",
        "Interdisziplinäre Betreuung verbessert Outcomes"
    ])

    # Slide 24: Thank you
    add_title_slide(prs,
        "Vielen Dank für Ihre Aufmerksamkeit!",
        "Fragen?")

    # Save the presentation
    prs.save('/home/user/Codex-stuff/Komplikationen_Magenbypass.pptx')
    print("Präsentation erfolgreich erstellt: Komplikationen_Magenbypass.pptx")

if __name__ == "__main__":
    create_presentation()
