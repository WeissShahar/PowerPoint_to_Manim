#%% Imports
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE, MSO_AUTO_SHAPE_TYPE
from manim import *
from manim_slides import Slide

# Initialize variables
presentation_path = 'rectangles.pptx'
slides_shapes_info = []  # Store shapes info for all slides
global_shapes = {}  # Global dictionary to store shapes info with shape_id as key
background_color = None

#%% Helper Functions
def convert_position(shape):
    """Convert PowerPoint coordinates to Manim coordinates."""
    return [
        shape.left.pt / 72 - config.frame_width / 2 + shape.width.pt / 72 / 2,
        config.frame_height / 2 - shape.top.pt / 72 - shape.height.pt / 72 / 2,
        0
    ]

def create_mobject(shape_info):
    """Create a Manim mobject based on the shape_info."""
    mobject = None
    if shape_info['type'] == 'rectangle':
        mobject = Rectangle(
            width=shape_info['dimensions'][0],
            height=shape_info['dimensions'][1]
        )
    elif shape_info['type'] == 'oval':
        diameter = min(shape_info['dimensions'])
        radius = diameter / 2
        mobject = Circle(radius=radius)
    elif shape_info['type'] == 'text':
        mobject = Text(shape_info['text'], font_size=24)
    
    if mobject:
        mobject.move_to(shape_info['position'])
    
    return mobject

def extract_shapes_from_slide(slide):
    """Function to extract relevant info from a slide and store shape details."""
    extracted_shapes = []
    for shape in slide.shapes:
        shape_info = {
            'id': shape.shape_id,
            'type': None,
            'position': convert_position(shape),
            'dimensions': (shape.width.pt / 72, shape.height.pt / 72),
            'text': shape.text if shape.has_text_frame and shape.text else None
        }

        if shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE:
            # Only supports Rectangle and Oval (circle) so far.
            if shape.auto_shape_type == MSO_AUTO_SHAPE_TYPE.RECTANGLE:
                shape_info['type'] = 'rectangle'
            elif shape.auto_shape_type == MSO_AUTO_SHAPE_TYPE.OVAL:
                shape_info['type'] = 'oval'
            else:
                continue
        elif shape.shape_type == MSO_SHAPE_TYPE.TEXT_BOX or shape.has_text_frame:
            if shape.has_text_frame and shape.text:
                shape_info['type'] = 'text'
            else:
                continue

        extracted_shapes.append(shape_info)
        global_shapes[shape_info['id']] = shape_info 

    return extracted_shapes


#%% Create and Animate Manim Objects
class PresentationScene(Slide):
    def __init__(self, slides_shapes_info, background_color=None, **kwargs):
        self.slides_shapes_info = slides_shapes_info
        self.background_color = background_color
        super().__init__(**kwargs)


    def initialize_first_slide(self, slide_shapes_info, mobjects):
        """Initialize shapes for the first slide."""
        for shape_info in slide_shapes_info:
            mobject = create_mobject(shape_info)
            if mobject:
                mobjects[shape_info['id']] = mobject
                self.add(mobject)
                self.wait(1)


    def construct(self):
        mobjects = {}

        self.initialize_first_slide(self.slides_shapes_info[0], mobjects) # Set the first slide

        for slide_number, shapes in enumerate(self.slides_shapes_info):
            if slide_number == 0:
                continue # Skip if it is the first slide
            
            current_shape_ids = {shape_info['id'] for shape_info in shapes} # Check for existing shapes

            for shape_id in list(mobjects.keys()):  # Remove mobjects that arent present
                if shape_id not in current_shape_ids:
                    self.remove(mobjects[shape_id])
                    del mobjects[shape_id]

            animations = []
            for shape_info in shapes: # Check for any changes in size / position of existing objects if exits. if not, create one
                shape_id = shape_info['id']
                mobject = mobjects.get(shape_id)

                if mobject:
                    new_position = shape_info['position']
                    new_dimensions = shape_info['dimensions']

                    if isinstance(mobject, Rectangle):
                        animations.append(mobject.animate.move_to(new_position).set(width=new_dimensions[0], height=new_dimensions[1]))
                    elif isinstance(mobject, Circle):
                        new_radius = min(new_dimensions) / 2
                        animations.append(mobject.animate.move_to(new_position).set(width=new_radius*2, height=new_radius*2))
                    elif isinstance(mobject, Text):
                        animations.append(mobject.animate.move_to(new_position))

                else:
                    mobject = create_mobject(shape_info)
                    if mobject:
                        mobjects[shape_id] = mobject
                        self.add(mobject)
                        self.wait(1)

            if animations:
                self.play(*animations)
                self.wait(1)

            # self.next_slide()

#%% Main Execution
if __name__ == "__main__":
    presentation = Presentation(presentation_path)
    
    for slide in presentation.slides:
        shapes = extract_shapes_from_slide(slide)
        slides_shapes_info.append(shapes)

    scene = PresentationScene(slides_shapes_info, background_color)
    
    # Render the scene
    config.media_width = "100%"
    scene.render()

# %%
