import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN

class PPTGenerator:
    # កំណត់ពណ៌ក្រុមហ៊ុន (Corporate Colors)
    COLORS = {
        'primary': RGBColor(0, 51, 102),      # Navy Blue
        'accent': RGBColor(245, 130, 32),     # Orange
        'text': RGBColor(33, 37, 41),         # Dark Grey
        'white': RGBColor(255, 255, 255),
        'light_blue': RGBColor(235, 245, 255),
        'gray': RGBColor(200, 200, 200),
        'green_excel': RGBColor(33, 115, 70),
        'grid_line': RGBColor(192, 192, 192),
        'trace_color': RGBColor(211, 211, 211)
    }

    def __init__(self):
        self.prs = Presentation()
        # កំណត់ទំហំស្លាយជា 16:9 (Widescreen)
        self.prs.slide_width = Inches(13.333)
        self.prs.slide_height = Inches(7.5)

    def set_font(self, run, size=18, is_title=False, color=None, is_bold=False, font_name=None):
        if font_name:
            run.font.name = font_name
        else:
            run.font.name = 'Khmer OS Moul Light' if is_title else 'Khmer OS Battambang'
        
        run.font.size = Pt(size)
        run.font.bold = is_bold
        if color:
            run.font.color.rgb = color

    def set_chinese_font(self, run, size=18, is_bold=True, color=None):
        run.font.name = 'Microsoft YaHei'
        run.font.size = Pt(size)
        run.font.bold = is_bold
        if color:
            run.font.color.rgb = color

    def add_header(self, slide, title_cn, title_km):
        """បន្ថែមរបារពណ៌ខៀវ និងចំណងជើងនៅគ្រប់ស្លាយ"""
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), self.prs.slide_width, Inches(1.2))
        bg.fill.solid()
        bg.fill.fore_color.rgb = self.COLORS['primary']
        bg.line.visible = False
        
        tb = slide.shapes.add_textbox(Inches(0.5), Inches(0.15), Inches(12), Inches(1))
        p = tb.text_frame.paragraphs[0]
        p.text = title_cn
        for run in p.runs: self.set_chinese_font(run, 28, True, self.COLORS['white'])
        
        p2 = tb.text_frame.add_paragraph()
        p2.text = title_km
        for run in p2.runs: self.set_font(run, 16, is_title=True, color=self.COLORS['white'])

    def create_cover(self):
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        # Background color
        rect = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, self.prs.slide_width, self.prs.slide_height)
        rect.fill.solid()
        rect.fill.fore_color.rgb = self.COLORS['light_blue']
        rect.line.visible = False

        # Central Box
        box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(3), Inches(2), Inches(7.333), Inches(3.5))
        box.fill.solid()
        box.fill.fore_color.rgb = self.COLORS['white']
        box.line.color.rgb = self.COLORS['primary']
        box.line.width = Pt(3)
        
        tb = slide.shapes.add_textbox(Inches(3.2), Inches(2.5), Inches(6.9), Inches(2.5))
        p = tb.text_frame.paragraphs[0]
        p.text = "第六课：成车异常与 Excel 计数"
        p.alignment = PP_ALIGN.CENTER
        for run in p.runs: self.set_chinese_font(run, 32, True, self.COLORS['primary'])
        
        p2 = tb.text_frame.add_paragraph()
        p2.text = "មេរៀនទី ៦៖ បញ្ហាកង់សម្រេច និង រូបមន្ត Excel (COUNTIF)"
        p2.alignment = PP_ALIGN.CENTER
        p2.space_before = Pt(20)
        for run in p2.runs: self.set_font(run, 20, is_title=True, color=self.COLORS['text'])
        
        p3 = tb.text_frame.add_paragraph()
        p3.text = "培训教师 : 郑和" 
        p3.alignment = PP_ALIGN.CENTER
        p3.space_before = Pt(30)
        for run in p3.runs: self.set_chinese_font(run, 16, True, self.COLORS['accent'])

    def create_vocab_slide(self, title_cn, title_km, vocab_list):
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        self.add_header(slide, title_cn, title_km)
        
        headers = ["中文", "拼音", "ភាសាខ្មែរ", "例句 (ឧទាហរណ៍)"]
        widths = [2.0, 2.2, 2.8, 5.5] 
        left = Inches(0.4)
        top = Inches(1.5)
        
        # គូរ Header តារាង
        current_x = left
        for h, w in zip(headers, widths):
            box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, current_x, top, Inches(w), Inches(0.6))
            box.fill.solid()
            box.fill.fore_color.rgb = self.COLORS['primary']
            box.line.color.rgb = self.COLORS['white']
            tb = slide.shapes.add_textbox(current_x, top, Inches(w), Inches(0.6))
            p = tb.text_frame.paragraphs[0]
            p.text = h
            p.alignment = PP_ALIGN.CENTER
            for run in p.runs: self.set_font(run, 14, is_title=True, color=self.COLORS['white'])
            current_x += Inches(w)

        # បំពេញទិន្នន័យ
        row_height = Inches(1.7)
        for idx, (cn, py, km, ex_cn, ex_km) in enumerate(vocab_list):
            y = top + Inches(0.7) + (row_height * idx) + (Inches(0.1 * idx))
            # Background row
            bg = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, y, sum([Inches(x) for x in widths]), row_height)
            bg.fill.solid()
            bg.fill.fore_color.rgb = self.COLORS['light_blue'] if idx % 2 == 0 else self.COLORS['white']
            bg.line.color.rgb = self.COLORS['gray']
            
            # Content (Chinese)
            tb = slide.shapes.add_textbox(left, y + Inches(0.5), Inches(widths[0]), Inches(0.6))
            p = tb.text_frame.paragraphs[0]
            p.text = cn; p.alignment = PP_ALIGN.CENTER
            for run in p.runs: self.set_chinese_font(run, 22, True, self.COLORS['primary'])

            # Pinyin
            tb = slide.shapes.add_textbox(left + Inches(widths[0]), y + Inches(0.6), Inches(widths[1]), Inches(0.6))
            p = tb.text_frame.paragraphs[0]
            p.text = py; p.alignment = PP_ALIGN.CENTER
            for run in p.runs: self.set_font(run, 15, font_name='Arial', color=self.COLORS['text'])

            # Khmer
            tb = slide.shapes.add_textbox(left + Inches(widths[0]+widths[1]), y + Inches(0.55), Inches(widths[2]), Inches(0.6))
            p = tb.text_frame.paragraphs[0]
            p.text = km; p.alignment = PP_ALIGN.CENTER
            for run in p.runs: self.set_font(run, 17, color=self.COLORS['text'])

            # Example
            tb = slide.shapes.add_textbox(left + Inches(sum(widths[:3])), y + Inches(0.2), Inches(widths[3]), Inches(1.3))
            p = tb.text_frame.paragraphs[0]
            p.text = ex_cn
            for run in p.runs: self.set_chinese_font(run, 14, False, self.COLORS['primary'])
            p2 = tb.text_frame.add_paragraph()
            p2.text = ex_km
            for run in p2.runs: self.set_font(run, 12, False, self.COLORS['text'])

    def create_excel_slide(self):
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        self.add_header(slide, "2. Excel 公式：计数 (COUNTIF)", "រូបមន្តរាប់ចំនួនតាមលក្ខខណ្ឌ")
        
        # Info Box
        box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(1.5), Inches(5.5), Inches(3))
        box.fill.solid()
        box.fill.fore_color.rgb = self.COLORS['light_blue']
        
        tb = slide.shapes.add_textbox(Inches(0.7), Inches(1.7), Inches(5.1), Inches(2.5))
        p = tb.text_frame.paragraphs[0]
        p.text = "🔢 COUNTIF Function"
        for run in p.runs: self.set_font(run, 24, True, self.COLORS['primary'], font_name='Arial')
        
        p2 = tb.text_frame.add_paragraph()
        p2.text = "រាប់ចំនួនក្រឡា (Cells) ដែលមានពាក្យដូចយើងចង់បាន។"
        p2.space_before = Pt(10)
        for run in p2.runs: self.set_font(run, 14, color=self.COLORS['text'])

        p3 = tb.text_frame.add_paragraph()
        p3.text = 'រូបមន្ត៖ =COUNTIF(Range, Criteria)'
        p3.space_before = Pt(20)
        for run in p3.runs: self.set_font(run, 16, True, self.COLORS['green_excel'], font_name='Consolas')

    def draw_tianzi_ge(self, slide, x, y, size, char=""):
        """គូរកងហាត់សរសេរ (Tianzi Ge)"""
        box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, size, size)
        box.fill.background()
        box.line.color.rgb = self.COLORS['primary']
        
        # គូរបន្ទាត់ចុចៗខាងក្នុង
        v = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, x + size/2, y, x + size/2, y + size)
        v.line.color.rgb = self.COLORS['grid_line']
        v.line.dash_style = 4
        h = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, x, y + size/2, x + size, y + size/2)
        h.line.color.rgb = self.COLORS['grid_line']
        h.line.dash_style = 4

        if char:
            tb = slide.shapes.add_textbox(x, y + Inches(0.05), size, size)
            p = tb.text_frame.paragraphs[0]; p.text = char; p.alignment = PP_ALIGN.CENTER
            for run in p.runs:
                run.font.name = 'Kaiti'
                run.font.size = Pt(38)
                run.font.color.rgb = self.COLORS['trace_color']

    def create_writing_practice(self, chars):
        """បង្កើតស្លាយហាត់សរសេរ"""
        chars_per_page = 7
        for i in range(0, len(chars), chars_per_page):
            slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
            self.add_header(slide, "附录：汉字练习", "ឧបសម្ព័ន្ធ៖ ការហាត់សរសេរអក្សរចិន")
            
            chunk = chars[i : i+chars_per_page]
            start_y = Inches(1.5)
            box_size = Inches(0.8)
            for idx, char in enumerate(chunk):
                curr_y = start_y + (idx * (box_size + Inches(0.05)))
                self.draw_tianzi_ge(slide, Inches(0.5), curr_y, box_size, char)
                for col in range(1, 14):
                    self.draw_tianzi_ge(slide, Inches(0.5) + (col * box_size), curr_y, box_size, "")

    def generate(self, output="QC_Lesson_06.pptx"):
        # បង្កើតស្លាយតាមលំដាប់លំដោយ
        self.create_cover()
        
        vocab_data = [
            [
                ("刹车失灵", "shā chē shī líng", "ហ្វ្រាំងមិនស៊ី", "后轮刹车失灵，需要维修。", "ហ្វ្រាំងក្រោយមិនស៊ីទេ ត្រូវការជួសជុល។"),
                ("变速不准", "biàn sù bù zhǔn", "ដូរលេខមិនចូល", "变速器不准，骑行不顺。", "ដូរលេខមិនត្រឹមត្រូវ ជិះមិនរលូនទេ។"),
                ("轮胎漏气", "lún tāi lòu qì", "សំបកកង់ធ្លាយ", "前门轮胎漏气了。", "សំបកកង់មុខធ្លាយខ្យល់ហើយ។")
            ],
            [
                ("螺丝松动", "luó sī sōng dòng", "ខ្ចៅធូរ", "脚踏螺丝松动，请锁紧。", "ខ្ចៅឈ្នាន់ធូរ សូមរឹតឱ្យតឹង។"),
                ("掉漆", "diào qī", "របកថ្នាំ", "架子掉漆，必须返工。", "តួកង់របកថ្នាំ ត្រូវតែធ្វើឡើងវិញ។"),
                ("划痕", "huá hén", "ស្នាមឆ្កូត", "包装前检查划痕。", "ត្រួតពិនិត្យស្នាមឆ្កូតមុនវេចខ្ចប់។")
            ]
        ]
        
        for i, group in enumerate(vocab_data):
            self.create_vocab_slide(f"1.{i+1} 常见异常", f"បញ្ហាទូទៅ {i+1}", group)
            
        self.create_excel_slide()
        
        all_chars = "刹车失灵变速不准轮胎漏气螺丝松动掉漆划痕"
        self.create_writing_practice(list(all_chars))
        
        self.prs.save(output)
        print(f"🚀 រួចរាល់! ឯកសារត្រូវបានរក្សាទុកជា: {output}")

if __name__ == "__main__":
    app = PPTGenerator()
    app.generate()