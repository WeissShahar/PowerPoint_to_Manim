from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE, MSO_AUTO_SHAPE_TYPE
from manim import *


class CreateShapesFromPPTX(Scene):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.shapes = []
        self.background_color = None

    def construct(self):
        presentation_path = 'rectangles_2.pptx'
        presentation = Presentation(presentation_path)
        slide = presentation.slides[0]

        self.extract_shapes_from_slide(slide)

        if self.background_color:
            self.camera.background_color = self.background_color

        for shape in self.shapes:
            self.add(shape)
            self.play(Create(shape))

        self.wait(1)

    def extract_shapes_from_slide(self, slide):
        """Function to extract relevant info from a slide and create Manim objects."""
        for shape in slide.shapes:
            if shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE:
                # Only supports Rectangle and Oval (circle) so far.
                if shape.auto_shape_type == MSO_AUTO_SHAPE_TYPE.RECTANGLE:
                    mobject = Rectangle(
                        width=shape.width.pt / 72,
                        height=shape.height.pt / 72
                    )
                elif shape.auto_shape_type == MSO_AUTO_SHAPE_TYPE.OVAL:
                    diameter = min(shape.width.pt, shape.height.pt) / 72
                    radius = diameter / 2
                    mobject = Circle(radius=radius)
                else:
                    continue

            # Adjust from Top-Left placement to Center placement
                mobject.move_to([
                    shape.left.pt / 72 - config.frame_width / 2 + shape.width.pt / 72 / 2,
                    config.frame_height / 2 - shape.top.pt / 72 - shape.height.pt / 72 / 2,
                    0
                ])
                                    # If shape includes a text frame
                if shape.has_text_frame and shape.text:
                    text_mobject = Text(shape.text, font_size=24)
                    text_mobject.move_to(mobject.get_center())
                    self.shapes.append(mobject)
                    self.shapes.append(text_mobject)
                else:
                    self.shapes.append(mobject)

                # If shape is text
            elif shape.shape_type == MSO_SHAPE_TYPE.TEXT_BOX or shape.has_text_frame:
                if shape.has_text_frame and shape.text:
                    text_mobject = Text(shape.text, font_size=24)
                    text_mobject.move_to([
                        shape.left.pt / 72 - config.frame_width / 2 + shape.width.pt / 72 / 2,
                        config.frame_height / 2 - shape.top.pt / 72 - shape.height.pt / 72 / 2,
                        0
                    ])
                    self.shapes.append(text_mobject)


if __name__ == "__main__":
    config.media_width = "100%"
    scene = CreateShapesFromPPTX()
    scene.render()
