from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

doc = Document()

# Page margins
section = doc.sections[0]
section.top_margin = Cm(2)
section.bottom_margin = Cm(2)
section.left_margin = Cm(2.5)
section.right_margin = Cm(2.5)

# Title
title = doc.add_heading('Solar Irradiation and Energy Data Sheet', level=1)
title.alignment = WD_ALIGN_PARAGRAPH.CENTER
title.runs[0].font.size = Pt(16)
title.runs[0].font.color.rgb = RGBColor(0, 0, 0)

# Subtitle
subtitle = doc.add_heading('Bogotá Savanna (Sabana de Bogotá), Colombia', level=2)
subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
subtitle.runs[0].font.size = Pt(13)
subtitle.runs[0].font.color.rgb = RGBColor(0, 0, 0)

doc.add_paragraph()

# Table data
rows = [
    ("Geographic Location", "4°35'N, 74°04'W", "—",
     "Located near the equator in the Eastern Cordillera of the Colombian Andes. Near-equatorial position ensures consistent solar angles and day lengths year-round."),
    ("Average Altitude", "~2,600", "m a.s.l.",
     "High altitude reduces atmospheric thickness, lowering air mass and increasing potential solar irradiance at the surface."),
    ("Climate Type", "Tropical Highland (Köppen: Cwb)", "—",
     "Mild temperatures year-round, bimodal rainfall pattern (two rainy and two dry seasons), and persistent cloud cover during rainy seasons."),
    ("Average Global Horizontal Irradiance (GHI)", "4.0 – 5.0", "kWh/m²/day",
     "Moderate-to-good irradiance for a tropical location. High altitude and equatorial position provide strong potential, partially reduced by cloudiness (~4.5 average)."),
    ("Annual GHI", "~1,600 – 1,700", "kWh/m²/year",
     "Derived from daily average (4.5 × 365 ≈ 1,643). Comparable to parts of southern Europe, making the region competitive for solar energy."),
    ("Direct Normal Irradiance (DNI)", "3.0 – 3.5", "kWh/m²/day",
     "Reduced by frequent cloud cover. This makes concentrating solar power (CSP) less viable, while flat-plate PV systems remain the preferred technology."),
    ("Diffuse Horizontal Irradiance (DHI)", "1.5 – 2.0", "kWh/m²/day",
     "High diffuse fraction (~35–45% of GHI) due to cloud scattering. Flat-plate PV panels capture diffuse radiation effectively from all sky directions."),
    ("Average Solar Peak Hours", "~4.5", "h/day",
     "Equivalent to daily GHI value. Represents hours of equivalent full irradiance (1,000 W/m²) per day — key parameter for PV system sizing."),
    ("Average Air Temperature", "13 – 14", "°C",
     "Cool and stable year-round due to altitude. Beneficial for PV: silicon panels lose ~0.4–0.5% efficiency per °C above 25°C, so cool conditions preserve rated output."),
    ("PV System Efficiency (Typical)", "15 – 20", "%",
     "Standard commercial monocrystalline or polycrystalline silicon panels. Cool operating temperatures help maintain efficiency close to rated values."),
    ("Estimated PV Energy Yield", "1,200 – 1,400", "kWh/kWp/year",
     "Calculated as: PSH × 365 × PR (0.75–0.80). Example: 4.5 × 365 × 0.78 ≈ 1,281 kWh/kWp/year. Sufficient for economically viable solar installations."),
    ("Seasonal Variability", "Low to Moderate", "—",
     "Bimodal rainfall causes ±15–20% GHI variation between seasons. Much lower variability than temperate regions due to consistent equatorial day length."),
    ("Main Limiting Factors", "Cloud cover, atmospheric humidity, seasonal rainfall", "—",
     "Persistent cloud cover in rainy seasons (Mar–May, Sep–Nov) is the primary limitation, reducing DNI significantly. Low dust accumulation is an advantage over arid regions."),
    ("Suitability for Solar Energy", "Moderate to Good", "—",
     "Equatorial location, high altitude, cool temperatures, and acceptable GHI make the Bogotá Savanna well-suited for flat-plate PV systems. Main engineering challenge is designing for high diffuse radiation and seasonal cloud variability."),
]

# Create table: 1 header row + 14 data rows, 4 columns
table = doc.add_table(rows=1 + len(rows), cols=4)
table.style = 'Table Grid'

# Set column widths
col_widths = [Cm(4.5), Cm(3.5), Cm(2.8), Cm(8.2)]
for i, width in enumerate(col_widths):
    for cell in table.columns[i].cells:
        cell.width = width

# Header row
headers = ["Parameter", "Value / Range", "Units", "Observations / Notes"]
header_row = table.rows[0]
for i, h in enumerate(headers):
    cell = header_row.cells[i]
    cell.text = h
    run = cell.paragraphs[0].runs[0]
    run.bold = True
    run.font.size = Pt(10)
    cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    # Gray background
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), 'BFBFBF')
    tcPr.append(shd)

# Data rows
for r_idx, (param, value, unit, note) in enumerate(rows):
    row = table.rows[r_idx + 1]
    data = [param, value, unit, note]
    for c_idx, text in enumerate(data):
        cell = row.cells[c_idx]
        cell.text = text
        para = cell.paragraphs[0]
        run = para.runs[0]
        run.font.size = Pt(9.5)
        if c_idx == 0:
            run.bold = True
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

# Conclusion paragraph
doc.add_paragraph()
conclusion = doc.add_paragraph()
conclusion.add_run('Conclusion: ').bold = True
conclusion.add_run(
    'The Bogotá Savanna is well-suited for photovoltaic solar energy. Its cool temperatures '
    'preserve panel efficiency, its near-equatorial position minimizes seasonal fluctuation, '
    'and its annual GHI is competitive globally. Flat-plate PV is the recommended technology '
    'given the high diffuse radiation fraction.'
)
conclusion.runs[-1].font.size = Pt(10)

output_path = '/home/user/Prognigg/Solar_Irradiation_Bogota_Savanna.docx'
doc.save(output_path)
print(f'Saved: {output_path}')
