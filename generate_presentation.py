from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

# إنشاء العرض التقديمي
prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
blank_layout = prs.slide_layouts[6]

# الألوان المعتمدة
DARK_BLUE = RGBColor(30, 58, 138)    # #1E3A8A (المتفاعلات)
DARK_GREEN = RGBColor(6, 95, 70)     # #065F46 (النواتج)
RED_COEFF = RGBColor(220, 38, 38)     # #DC2626 (المعاملات)
GRAY_SUB = RGBColor(51, 65, 85)      # #334155 (الأرقام السفلية)
LIGHT_BG = RGBColor(250, 248, 245)

def add_header(slide, title_text):
    # شريط العنوان
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(1.1))
    shape.fill.solid()
    shape.fill.fore_color.rgb = DARK_BLUE
    shape.line.fill.background()
    
    tf = shape.text_frame
    tf.margin_top = Inches(0.2)
    p = tf.paragraphs[0]
    p.text = title_text
    p.font.size = Pt(26)
    p.font.bold = True
    p.font.color.rgb = RGBColor(255, 255, 255)
    p.alignment = PP_ALIGN.CENTER

# ----------------------------------------------------
# الشريحة 1: شريحة العنوان
# ----------------------------------------------------
slide1 = prs.slides.add_slide(blank_layout)
bg1 = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
bg1.fill.solid()
bg1.fill.fore_color.rgb = DARK_BLUE
bg1.line.fill.background()

txBox = slide1.shapes.add_textbox(Inches(1), Inches(1.8), Inches(11.333), Inches(4))
tf = txBox.text_frame
tf.word_wrap = True

p = tf.paragraphs[0]
p.text = "درس: وزن المعادلات الكيميائية"
p.font.size = Pt(40)
p.font.bold = True
p.font.color.rgb = RGBColor(255, 255, 255)
p.alignment = PP_ALIGN.CENTER

p2 = tf.add_paragraph()
p2.text = "الصف الأول الثانوي - المنهج السوداني (الوحدة الثالثة)"
p2.font.size = Pt(22)
p2.font.color.rgb = RGBColor(147, 197, 253)
p2.alignment = PP_ALIGN.CENTER

p3 = tf.add_paragraph()
p3.text = "\nإعداد وتجميع المادة العلمية: د. رشا يعقوب عبدالله\nزمن الحصة: 45 دقيقة"
p3.font.size = Pt(18)
p3.font.color.rgb = RGBColor(226, 232, 240)
p3.alignment = PP_ALIGN.CENTER

# ----------------------------------------------------
# الشريحة 2: التمهيد والمفاهيم الذهبية
# ----------------------------------------------------
slide2 = prs.slides.add_slide(blank_layout)
add_header(slide2, "1. التمهيد والمفاهيم الأساسية (20 دقيقة)")

txBox = slide2.shapes.add_textbox(Inches(0.8), Inches(1.5), Inches(11.733), Inches(5.5))
tf = txBox.text_frame
tf.word_wrap = True

p = tf.paragraphs[0]
p.text = "• قانون بقاء المادة (حفظ الكتلة):"
p.font.size = Pt(20)
p.font.bold = True
p.font.color.rgb = DARK_BLUE

p = tf.add_paragraph()
p.text = "   \"المادة لا تفنى ولا تستحدث في التفاعل الكيميائي، بل تعيد الذرات ترتيب نفسها فقط.\""
p.font.size = Pt(18)

p = tf.add_paragraph()
p.text = "\n• المفاهيم الرئيسية:"
p.font.size = Pt(20)
p.font.bold = True
p.font.color.rgb = DARK_BLUE

p = tf.add_paragraph()
p.text = "   1. المعادلة الموزونة: تتساوى فيها أعداد ذرات العناصر في طرفي التفاعل."
p.font.size = Pt(18)

p = tf.add_paragraph()
p.text = "   2. المعامل (Coefficient): الرقم الذي يكتب يسار الصيغة ويحدد عدد الجزيئات (قابل للتعديل)."
p.font.size = Pt(18)

p = tf.add_paragraph()
p.text = "   3. الرقم السفلي (Subscript): الرقم أسفل يمين الرمز ويحدد تركيب المركب (يُحظر تعديله تماماً)."
p.font.size = Pt(18)

# ----------------------------------------------------
# الشريحة 3: أمثلة وزن المعادلات
# ----------------------------------------------------
slide3 = prs.slides.add_slide(blank_layout)
add_header(slide3, "2. أمثلة تطبيقية لوزن المعادلات (15 دقيقة)")

txBox = slide3.shapes.add_textbox(Inches(0.8), Inches(1.4), Inches(11.733), Inches(5.5))
tf = txBox.text_frame
tf.word_wrap = True

p = tf.paragraphs[0]
p.text = "مثال (1): تفاعل النيتروجين والهيدروجين لتكوين النشادر"
p.font.size = Pt(20)
p.font.bold = True
p.font.color.rgb = DARK_BLUE

p = tf.add_paragraph()
p.text = "  غير موزونة: N2 + H2 -> NH3"
p.font.size = Pt(18)

p = tf.add_paragraph()
p.text = "  المعادلـة الموزونــة: N2 + 3H2 -> 2NH3"
p.font.size = Pt(22)
p.font.bold = True
p.font.color.rgb = DARK_GREEN

p = tf.add_paragraph()
p.text = "\nمثال (2): احتراق غاز الميثان"
p.font.size = Pt(20)
p.font.bold = True
p.font.color.rgb = DARK_BLUE

p = tf.add_paragraph()
p.text = "  غير موزونة: CH4 + O2 -> CO2 + H2O"
p.font.size = Pt(18)

p = tf.add_paragraph()
p.text = "  المعادلـة الموزونــة: CH4 + 2O2 -> CO2 + 2H2O"
p.font.size = Pt(22)
p.font.bold = True
p.font.color.rgb = DARK_GREEN

# ----------------------------------------------------
# الشريحة 4: التقويم والواجب
# ----------------------------------------------------
slide4 = prs.slides.add_slide(blank_layout)
add_header(slide4, "3. التقويم الصفي والواجب المنزلي (10 دقائق)")

txBox = slide4.shapes.add_textbox(Inches(0.8), Inches(1.5), Inches(11.733), Inches(5.5))
tf = txBox.text_frame
tf.word_wrap = True

p = tf.paragraphs[0]
p.text = "• التطبيق الصفي (تمرين سريع):"
p.font.size = Pt(20)
p.font.bold = True
p.font.color.rgb = DARK_BLUE

p = tf.add_paragraph()
p.text = "  1. Mg + 2HCl -> MgCl2 + H2"
p.font.size = Pt(18)
p = tf.add_paragraph()
p.text = "  2. Fe + 2HCl -> FeCl2 + H2"
p.font.size = Pt(18)

p = tf.add_paragraph()
p.text = "\n• الواجب المنزلي:"
p.font.size = Pt(20)
p.font.bold = True
p.font.color.rgb = RED_COEFF

p = tf.add_paragraph()
p.text = "  حوّل المعادلة اللفظية إلى رمزية موزونة:\n  \"يتفاعل الكالسيوم مع حمض الكبريتيك لإنتاج كبريتات الكالسيوم وينطلق غاز الهيدروجين\""
p.font.size = Pt(18)

# حفظ الملف
prs.save("balancing_chemical_equations.pptx")
print("PowerPoint presentation generated successfully!")
print("File saved as: balancing_chemical_equations.pptx")
