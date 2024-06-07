from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE, MSO_AUTO_SHAPE_TYPE
from manim import *


def extract_shapes_from_slide(slide):
    """Function to extract relevant info from a slide."""
    shapes_info = []
    for shape in slide.shapes:
        shape_info = {
            "type": shape.shape_type,
            "name": shape.name
        }

        if shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE:
            shape_info["auto_shape_type"] = shape.auto_shape_type # Provides info about the type of shape used

        shapes_info.append(shape_info)
    return shapes_info



if __name__ == "__main__":
    presentation_path = 'Simple Example.pptx' # A simple one slide presentation created for the basic pptx -> Manim
    presentation = Presentation(presentation_path)

    slide = presentation.slides[0]
    shapes_info = extract_shapes_from_slide(slide)
    print(shapes_info)
