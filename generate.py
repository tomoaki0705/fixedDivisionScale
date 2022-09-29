from cmath import log
from pptx.enum.shapes import MSO_CONNECTOR
from pptx import Presentation
from pptx.util import Cm
import math

prs = Presentation('A4_template.pptx')
title_slide_layout = prs.slide_layouts[6]
slide = prs.slides.add_slide(title_slide_layout)

# title.text = "Hello, World!"
# subtitle.text = "python-pptx was here!"

def drawLogLine(begin,inclusiveEnd,base,slide,y,height=1,left=1,right=26):
    horizontalLength=right-left
    scale=horizontalLength/(math.log(inclusiveEnd)/math.log(base))
    for i in range(begin,inclusiveEnd+1):
        position=scale*(math.log(i)/math.log(10))+left
        slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Cm(position), Cm(y), Cm(position), Cm(y+1))

# line1=slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, 1, Cm(2), Cm(1), Cm(2))
# line1=slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Cm(1), Cm(1), Cm(3), Cm(3))

drawLogLine(1,10,10,slide,3)

prs.save('test.pptx')