from pptx import Presentation
from pptx.util import Inches
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.dml import MSO_THEME_COLOR


def add_bullet_slide(pr, title, bullet_points):
    """Add a bullet slide to *pr* with *title* and bullet text."""
    slide_layout = pr.slide_layouts[1]
    slide = pr.slides.add_slide(slide_layout)
    slide.shapes.title.text = title
    body_shape = slide.shapes.placeholders[1]
    if bullet_points:
        body_shape.text = bullet_points[0]
        tf = body_shape.text_frame
        for text in bullet_points[1:]:
            p = tf.add_paragraph()
            p.text = text
    return slide
pr1 = Presentation()

slide1_layout = pr1.slide_layouts[0]

slide1 = pr1.slides.add_slide(slide1_layout)

title1 = slide1.shapes.title
subtitle1 = slide1.placeholders[1]

title1.text= "ANALYSTRISING"
subtitle1.text = "Subscribe to my channel"

add_bullet_slide(pr1, "Now For Some Bullet Points", ["Subscribe", "to", "my", "Channel!"])

#Add Slide 3
slide3_layout = pr1.slide_layouts[5]
slide3 = pr1.slides.add_slide(slide3_layout)
title3 = slide3.shapes.title
title3.text = "Picture Time!"

img1 = "Elements.jpg"
from_left = Inches(3)
from_top = Inches(4)
add_picture = slide3.shapes.add_picture(img1,from_left,from_top)

#Part3
#Add Slide 4
slide4_layout = pr1.slide_layouts[5]
slide4 = pr1.slides.add_slide(slide4_layout)
title4 = slide4.shapes.title
title4.text = "Shapework"

#Create a shape
left1 = top1 = width1 = height1 = Inches(2)
add_shape1 = slide4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,left1,top1,width1,height1)

left2 = Inches(6)
top2 = Inches(2)
width2 = height2 = Inches(2)
arrow1 = slide4.shapes.add_shape(MSO_SHAPE.DOWN_ARROW,left2,top2,width2,height2)


fill_arrow1 = arrow1.fill
fill_arrow1.solid()
fill_arrow1.fore_color.theme_color = MSO_THEME_COLOR.ACCENT_5
arrow1.text = "Pijl111"

arrow1.rotation = 180

#pr1.save('GreatPresentation.pptx')
#pr1.save('GreatPresentation_Part2.pptx')
pr1.save('GreatPresentation_Part3.pptx')


