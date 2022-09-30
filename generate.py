from cmath import log
from pptx.enum.shapes import MSO_CONNECTOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.dml import MSO_FILL_TYPE
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.dml.line import LineFormat
from pptx.shapes.connector import Connector
from pptx import Presentation
from pptx.util import Cm, Pt
import math

prs = Presentation('A4_template.pptx')
title_slide_layout = prs.slide_layouts[6]
slide = prs.slides.add_slide(title_slide_layout)

# title.text = "Hello, World!"
# subtitle.text = "python-pptx was here!"

def drawLog10Line(begin,inclusiveEnd,slide,y,height=1,left=1,right=26,indexScale=1):
    horizontalLength=right-left
    scale=horizontalLength/(math.log10(inclusiveEnd))
    slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Cm(left), Cm(y), Cm(right), Cm(y))
    for i in range(begin,inclusiveEnd+1):
        scaledIndex = i * indexScale
        # print(f"scaledIndex:{scaledIndex}")
        position=scale*(math.log10(scaledIndex/(begin*indexScale)))+left
        slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Cm(position), Cm(y-0.5), Cm(position), Cm(y+0.5))
        # print(f"position:{position}")
        textBox = slide.shapes.add_textbox(Cm(position-0.5), Cm(y-2), Cm(1), Cm(1))
        paragraph0 = textBox.text_frame
        paragraph0.text = str(scaledIndex)
    for i in range(begin,inclusiveEnd):
        scaledIndex = i * indexScale
        for j in range(0,10):
            fineTics = scaledIndex + 0.1 * j
            position=scale*(math.log10(fineTics/(begin*indexScale)))+left
            slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Cm(position), Cm(y-0.2), Cm(position), Cm(y+0.2))


def patch_connector():
    def get_or_add_ln(self):
        return self._element.spPr.get_or_add_ln()
    Connector.get_or_add_ln = get_or_add_ln
patch_connector()

def drawLog10LineInvert(begin,inclusiveEnd,slide,y,height=1,left=1,right=26,indexScale=1):
    horizontalLength=right-left
    scale=horizontalLength/(math.log10(inclusiveEnd))
    slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Cm(left), Cm(y), Cm(right), Cm(y))
    for i in range(begin,inclusiveEnd+1):
        scaledIndex = i * indexScale
        # print(f"scaledIndex:{scaledIndex}")
        position=right - scale*(math.log10(scaledIndex/(begin*indexScale)))
        line = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Cm(position), Cm(y-0.5), Cm(position), Cm(y+0.5))
        # line.ln = line.get_or_add_ln
        # lineFormat = LineFormat(line)
        # lineFormat.fill.fore_color.rgb = RGBColor(255, 0, 0)
        # line.color.rgb = RGBColor(255, 0, 0)
        # print(f"position:{position}")
        textBox = slide.shapes.add_textbox(Cm(position-0.5), Cm(y-2), Cm(1), Cm(1))
        paragraph0 = textBox.text_frame
        paragraph0.text = str(scaledIndex)
        # paragraph0.font.color.rgb = RGBColor(255,0,0)
    for i in range(begin,inclusiveEnd):
        scaledIndex = i * indexScale
        for j in range(0,10):
            fineTics = scaledIndex + 0.1 * j
            position=right - scale*(math.log10(fineTics/(begin*indexScale)))
            slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Cm(position), Cm(y-0.2), Cm(position), Cm(y+0.2))

def drawLog2Line(begin,end,ticNumber,slide,y,height=1,left=1,right=26,indexScale=1):
    horizontalLength=right-left
    scale=horizontalLength/(math.log2(end/begin))
    slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Cm(left), Cm(y), Cm(right), Cm(y))
    # print (f"bgein:{begin}")
    # print (f"end  :{end}")
    leftLog = math.log2(begin)
    # print(f"leftLog    :{leftLog}")
    for i in range(0,ticNumber):
        scaledIndex = i * indexScale + begin
        position=scale*((math.log2(scaledIndex))-leftLog)+left
        # print(f"scaledIndex:{scaledIndex}")
        # print(f"position   :{position}")
        slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Cm(position), Cm(y-0.5), Cm(position), Cm(y+0.5))
        textBox = slide.shapes.add_textbox(Cm(position-0.5), Cm(y-1), Cm(1), Cm(1))
        textFrame0 = textBox.text_frame
        paragraph0 = textFrame0.paragraphs[0]
        paragraph0.text = str(scaledIndex)
        paragraph0.font.size = Pt(8)
        paragraph0.alignment = PP_ALIGN.CENTER
    for i in range(0,ticNumber-1):
        scaledIndex = i * indexScale + begin
        for j in range(1,8):
            fineTics = (scaledIndex + (j * begin/16)) * 1.0
            position=scale*(math.log2(fineTics)-leftLog)+left
            # print(f"{fineTics},{position}")
            slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Cm(position), Cm(y-0.2), Cm(position), Cm(y+0.2))

# line1=slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, 1, Cm(2), Cm(1), Cm(2))
# line1=slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Cm(1), Cm(1), Cm(3), Cm(3))

verticalOffset = 2.5
offset = verticalOffset
drawLog10Line(1,10,slide,offset)
offset += verticalOffset
# drawLogLine(1,2,11,2,slide,offset,indexScale=0.1)
# offset += verticalOffset
drawLog2Line(16,160,19,slide,offset,indexScale=8)
offset += verticalOffset
drawLog2Line(160,1600,19,slide,offset,indexScale=80)
offset += verticalOffset
# drawLogLine(160,1600,46,2,slide,offset,indexScale=32)
# offset += verticalOffset
drawLog2Line(32,320,19,slide,offset,indexScale=16)
offset += verticalOffset
drawLog2Line(64,640,19,slide,offset,indexScale=32)
offset += verticalOffset
drawLog2Line(128,1280,19,slide,offset,indexScale=64)
offset += verticalOffset

title_slide_layout = prs.slide_layouts[6]
slide2 = prs.slides.add_slide(title_slide_layout)

offset = verticalOffset
drawLog10LineInvert(1,10,slide2,offset)
divScale=(26-1)
divLeft=26-(divScale * math.log10(8))
offset += verticalOffset
drawLog2Line(2,16,15,slide2,offset,indexScale=1,left=divLeft)
drawLog2Line(1.6,2,1,slide2,offset,indexScale=0.1,right=divLeft)
offset += verticalOffset
drawLog2Line(4,32,28,slide2,offset,indexScale=1,left=divLeft)
drawLog2Line(3.2,4,1,slide2,offset,indexScale=0.1,right=divLeft)
offset += verticalOffset
divLeft=26-(divScale * math.log10(64/7))
drawLog2Line(7,64,58,slide2,offset,indexScale=1,left=divLeft)
drawLog2Line(6.4,7,1,slide2,offset,indexScale=0.1,right=divLeft)


prs.save('test.pptx')