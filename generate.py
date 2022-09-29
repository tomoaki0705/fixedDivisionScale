from cmath import log
from pptx.enum.shapes import MSO_CONNECTOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.dml import MSO_FILL_TYPE
from pptx.dml.color import RGBColor
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
        # print(f"position:{position}")
        textBox = slide.shapes.add_textbox(Cm(position-0.5), Cm(y-2), Cm(1), Cm(1))    #Text Box Shapeオブジェクトの追加
        paragraph0 = textBox.text_frame
        paragraph0.text = str(scaledIndex)
        slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Cm(position), Cm(y-0.5), Cm(position), Cm(y+0.5))

def drawLogLine(begin,end,ticNumber,base,slide,y,height=1,left=1,right=26,indexScale=1):
    horizontalLength=right-left
    scale=horizontalLength/(math.log(end)/math.log(base))
    slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Cm(left), Cm(y), Cm(right), Cm(y))
    print (f"bgein:{begin}")
    print (f"end  :{end}")
    leftLog = math.log(begin)/math.log(base)
    for i in range(0,ticNumber):
        scaledIndex = i * indexScale + begin
        print(f"scaledIndex:{scaledIndex}")
        position=scale*(math.log(scaledIndex)/math.log(base))-leftLog+left
        print(f"position:{position}")
        slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Cm(position), Cm(y-0.5), Cm(position), Cm(y+0.5))

# line1=slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, 1, Cm(2), Cm(1), Cm(2))
# line1=slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Cm(1), Cm(1), Cm(3), Cm(3))

drawLog10Line(1,10,slide,3)
drawLogLine(1,2,11,2,slide,5,indexScale=0.1)

prs.save('test.pptx')