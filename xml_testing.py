from pptx import Presentation
import xml.etree.ElementTree as ET
from lxml import etree
from manim import *
from manim_slides import Slide
background_color = WHITE
# Function to parse the theme XML and create a map of theme colors
def parse_theme_colors(theme_xml):
    ns = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
    root = ET.fromstring(theme_xml)
    theme_colors = {}

    # Find the theme color definitions in the clrScheme tag
    color_schemes = root.findall('.//a:clrScheme', namespaces=ns)

    if color_schemes:
        for scheme in color_schemes:
            for color in scheme:
                color_tag = color.tag.split('}')[1]  # Extract the tag after namespace

                # Handle <a:srgbClr>
                srgb_clr = color.find('.//a:srgbClr', namespaces=ns)
                if srgb_clr is not None:
                    color_val = srgb_clr.attrib.get('val')  # Example: "FFFFFF"
                    theme_colors[color_tag] = f'#{color_val}'
                    continue  # Skip to the next color if found

                # Handle <a:sysClr> (system colors like bg1, tx1)
                sys_clr = color.find('.//a:sysClr', namespaces=ns)
                if sys_clr is not None:
                    color_val = sys_clr.attrib.get('lastClr')  # Example: "FFFFFF"
                    theme_colors[color_tag] = f'#{color_val}'
                    continue  # Skip to the next color if found

                # Handle luminance modification (lumMod and lumOff)
                scheme_clr = color.find('.//a:schemeClr', namespaces=ns)
                if scheme_clr is not None:
                    lum_mod = scheme_clr.find('.//a:lumMod', namespaces=ns)
                    lum_off = scheme_clr.find('.//a:lumOff', namespaces=ns)
                    if lum_mod is not None or lum_off is not None:
                        theme_colors[color_tag] = f'Luminance Modified Color'
                    else:
                        theme_colors[color_tag] = 'Scheme Color'

    return theme_colors

# Function to get the adjusted color with luminance and shade applied
def get_adjusted_color(color_val, lum_mod=None, shade_mod=None, theme_colors=None):
    if theme_colors:
        base_color = theme_colors.get(color_val, '#FFFFFF')  # Default to white if not found

        # Convert hex to RGB
        rgb = [int(base_color[i:i+2], 16) for i in (1, 3, 5)]

        # Apply luminance adjustment if present
        if lum_mod is not None:
            lum_factor = lum_mod / 100000
            rgb = [int(c * lum_factor) for c in rgb]

        # Apply shade adjustment if present
        if shade_mod is not None:
            shade_factor = shade_mod / 100000
            rgb = [int(c * (1 + shade_factor)) for c in rgb]

        # Ensure RGB values are within valid range
        rgb = [max(0, min(255, c)) for c in rgb]

        # Convert back to hex
        return '#{:02X}{:02X}{:02X}'.format(*rgb)
    return '#FFFFFF'


def extract_shape_info(slide_xml, master_slide_clr_mapping, theme_colors):
    ns = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
          'p': 'http://schemas.openxmlformats.org/presentationml/2006/main'}

    shape_info = []

    root = etree.fromstring(slide_xml)
    shapes = root.findall('.//p:sp', namespaces=ns)

    # Find all shapes
    for shape in shapes:
        info = {}

        # Extract shape type
        shape_type_elem = shape.find('.//p:nvSpPr/p:cNvPr', namespaces=ns)
        if shape_type_elem is not None:
            shape_name = shape_type_elem.get('name')
            shape_id = shape_type_elem.get('id')  # Extract shape ID
            if shape_id:
                info['id'] = shape_id  # Save the ID
            if shape_name:
                if 'Oval' in shape_name:
                    info['type'] = 'Oval'
                else:
                    info['type'] = 'Unknown'

        # Extract position and size
        xfrm = shape.find('.//a:xfrm', namespaces=ns)
        if xfrm is not None:
            off = xfrm.find('a:off', namespaces=ns)
            ext = xfrm.find('a:ext', namespaces=ns)
            if off is not None and ext is not None:
                info['position'] = {'x': int(off.get('x')), 'y': int(off.get('y'))}
                info['size'] = {
                    'cx': int(ext.get('cx')) / 914400,  # Convert from EMU to points
                    'cy': int(ext.get('cy')) / 914400  # Convert from EMU to points
                }
        else:
            info['position'] = None
            info['size'] = None

        # Extract shape color
        solid_fill = shape.find('.//a:solidFill', namespaces=ns)
        if solid_fill is not None:
            scheme_clr = solid_fill.find('a:schemeClr', namespaces=ns)
            if scheme_clr is not None:
                color_val = scheme_clr.get('val')
                # Map to actual color using master_slide_clr_mapping
                actual_color_val = master_slide_clr_mapping.get(color_val, color_val)
                # Get the actual color value from theme_colors
                info['color'] = get_adjusted_color(
                    actual_color_val,
                    lum_mod=int(scheme_clr.find('a:lumMod', namespaces=ns).get('val')) if scheme_clr.find('a:lumMod',
                                                                                                          namespaces=ns) is not None else None,
                    shade_mod=int(scheme_clr.find('a:shade', namespaces=ns).get('val')) if scheme_clr.find('a:shade',
                                                                                                           namespaces=ns) is not None else None,
                    theme_colors=theme_colors
                )
            else:
                info['color'] = None
        else:
            info['color'] = None

        shape_info.append(info)

        # Extract text content and treat it as a separate Text object
        text_elems = shape.findall('.//a:t', namespaces=ns)
        if text_elems:
            text_info = {}

            text_content = '\n'.join([elem.text for elem in text_elems if elem.text is not None])
            text_info['type'] = 'Text'
            text_info['text'] = text_content

            # Use the same position and size as the shape for the text
            text_info['position'] = info.get('position')
            text_info['size'] = info.get('size')

            # Extract text color (using the same color field)
            rPr = shape.find('.//a:rPr', namespaces=ns)
            if rPr is not None:
                text_fill = rPr.find('.//a:solidFill/a:schemeClr', namespaces=ns)
                if text_fill is not None:
                    text_color_val = text_fill.get('val')
                    # Map to actual color using master_slide_clr_mapping
                    actual_text_color_val = master_slide_clr_mapping.get(text_color_val, text_color_val)
                    # Get the actual text color value from theme_colors
                    text_info['color'] = get_adjusted_color(
                        actual_text_color_val,
                        lum_mod=int(text_fill.find('a:lumMod', namespaces=ns).get('val')) if text_fill.find('a:lumMod',
                                                                                                            namespaces=ns) is not None else None,
                        shade_mod=int(text_fill.find('a:shade', namespaces=ns).get('val')) if text_fill.find('a:shade',
                                                                                                             namespaces=ns) is not None else None,
                        theme_colors=theme_colors
                    )
                else:
                    text_info['color'] = '#000000'  # Default text color if not found
            else:
                text_info['color'] = '#000000'  # Default text color if no formatting is found

            # Extract font size
            if rPr is not None:
                sz = rPr.get('sz')
                text_info['font_size'] = int(sz) / 100 if sz is not None else 27  # Default font size if not found
            else:
                text_info['font_size'] = 27  # Default font size if no formatting is found

            # Append the text object to the list
            shape_info.append(text_info)

    return shape_info



EMU_TO_POINTS = 1 / 12700


def convert_position(shape):
    """Convert PowerPoint coordinates to Manim coordinates."""

    # Convert from EMU to points
    x_in_points = shape['position']['x'] * EMU_TO_POINTS
    y_in_points = shape['position']['y'] * EMU_TO_POINTS

    # Convert to Manim coordinates
    manim_x = x_in_points / 72 - config.frame_width / 2 + shape['size']['cx'] / 2
    manim_y = config.frame_height / 2 - (y_in_points / 72 + shape['size']['cy'] / 2)

    return [manim_x, manim_y, 0]


def extract_color_values_from_theme(theme_xml):
    root = ET.fromstring(theme_xml)
    ns = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
    color_values = {}

    for color in root.findall('.//a:clrScheme/*', ns):
        color_name = color.tag.split('}')[-1]  # Get the tag name without namespace
        color_value = color.find('.//a:srgbClr', ns)
        if color_value is not None:
            color_values[color_name] = color_value.get('val')
        else:
            color_value = color.find('.//a:sysClr', ns)
            if color_value is not None:
                color_values[color_name] = color_value.get('lastClr')

    return color_values

def map_colors_to_values(master_slide_clr_mapping, color_values):
    final_color_mapping = {}

    for master_color, theme_color in master_slide_clr_mapping.items():
        final_color_mapping[master_color] = color_values.get(theme_color, 'Color not found')

    return final_color_mapping

def extract_master_slide_clr_mapping(master_slide_xml):
    root = ET.fromstring(master_slide_xml)
    ns = {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main'}
    clr_map_elem = root.find('p:clrMap', ns)

    if clr_map_elem is not None:
        return {key: value for key, value in clr_map_elem.attrib.items()}
    else:
        return {}

def create_mobject(shape_info):
    """Create a Manim mobject based on the shape_info."""
    mobject = None
    if shape_info['type'] == 'rectangle':
        mobject = Rectangle(
            width=shape_info['dimensions'][0],
            height=shape_info['dimensions'][1],
            color=BLACK
        )
    elif shape_info['type'] == 'Oval':
        mobject = Ellipse(
            width=shape_info['size']['cx'],
            height=shape_info['size']['cy'],
        )

        # Apply the color if it exists
        if 'color' in shape_info:
            mobject.set_fill(color=shape_info['color'], opacity=1)
            mobject.set_stroke(color=BLACK)  # You can adjust the width as needed



    elif shape_info['type'] == 'Text':
        if shape_info['text'] is None:
            shape_info['text']='failed'
        font_size = shape_info.get('font_size', 27)  # Default to 27 if not provided

        text_color = shape_info.get('text_color', {'val': '#000000'})  # Default to black if not provided

        # Ensure text_color is not None and has a 'val' key

        if text_color is None:
            text_color = {'val': '#000000'}

        if font_size is None:
            font_size = 24
        mobject = Text(

            shape_info['text'],

            font_size=font_size,

            color=text_color['val']

        )
    elif shape_info['type'] == 'image' and shape_info['image_path']:
        mobject = ImageMobject(shape_info['image_path'])
        mobject.width, mobject.height = shape_info['dimensions']
    elif shape_info['type'] == 'Line':
        start_point, end_point = shape_info['dimensions']
        print(shape_info['color'])
        if shape_info['dash_style'] == 'dashed':
            mobject = DashedLine(start=start_point, end=end_point, color= ManimColor.from_rgb(shape_info['color']))
        else:
            mobject = Line(start=start_point, end=end_point, color= ManimColor.from_rgb(shape_info['color']))
    elif shape_info['type'] == 'arrow':
        start_point, end_point = shape_info['dimensions']
        mobject = Arrow(start=start_point, end=end_point, color=BLACK, stroke_width=1, max_stroke_width_to_length_ratio=1, max_tip_length_to_length_ratio=0.1)
    elif shape_info['type'] == 'table':
        table_data = shape_info['table_data']
        mobject = MathTable(
            table_data,
            include_outer_lines=True,
            line_config={"stroke_color": BLACK, "stroke_width": 2},
            element_to_mobject_config={"color": BLACK},
        )
        mobject.scale(0.2)  # Scale down the table to fit the scene

    if mobject and shape_info['position']:
        mobject.move_to(shape_info['position'])

    return mobject
class PresentationScene(Slide):
    def __init__(self, slides_shapes_info, background_color=None, **kwargs):
        self.slides_shapes_info = slides_shapes_info
        super().__init__(**kwargs)

    def initialize_slide(self, slide_shapes_info):
        """Initialize shapes for a slide."""
        for shape_info in slide_shapes_info:
            mobject = create_mobject(shape_info)
            if mobject:
                self.add(mobject)

    def construct(self):
        self.camera.background_color = background_color
        for slide_shapes_info in self.slides_shapes_info:
            self.clear()  # Clear all shapes from the previous slide
            self.initialize_slide(slide_shapes_info)
            self.wait(1)
            self.next_slide()
def main():
    # Load the presentation
    presentation = Presentation('presentations/PathPlanning.pptx')

    # Get the theme part from the related parts of the presentation
    theme_part = None
    for rel in presentation.part.rels.values():
        if "theme" in rel.target_ref:
            theme_part = rel.target_part
            break

    # Get the theme XML (if any)
    theme_xml = theme_part._blob.decode() if theme_part else None

    # Extract color values from the theme XML
    color_values = extract_color_values_from_theme(theme_xml) if theme_xml else {}

    # Get the Master Slide XML
    master_slide = presentation.slide_master
    master_slide_xml = master_slide._element.xml
    master_slide_clr_mapping = extract_master_slide_clr_mapping(master_slide_xml)

    # Map master slide colors to actual color values
    final_color_mapping = map_colors_to_values(master_slide_clr_mapping, color_values)

    # Extract color values from theme
    theme_colors = parse_theme_colors(theme_xml) if theme_xml else {}

    # Extract shape info
    slides_shapes_info = []  # Store shapes info for all slides
    # Get slide width and height in inches
    slide_width_in_inches = presentation.slide_width.inches
    slide_height_in_inches = presentation.slide_height.inches

    # Set Manim frame size based on PowerPoint slide dimensions
    config.frame_width = slide_width_in_inches  # Convert to points (1 inch = 72 points)
    config.frame_height = slide_height_in_inches  # Convert to points (1 inch = 72 points)

    # Optionally set Manim's pixel resolution to match PowerPoint's
    config.pixel_width = int(slide_width_in_inches * config.pixel_height / slide_height_in_inches)
    for slide in presentation.slides:
        slide_xml = slide._element.xml

        # Extract shape info for the current slide
        slide_shape_info = extract_shape_info(slide_xml, master_slide_clr_mapping, theme_colors)

        # Convert positions for shapes on the slide


        for shape_info in slide_shape_info:
            if shape_info.get('position'):
                shape_info['position'] = convert_position(
                    shape_info,
                )

        # Append the shape info for the current slide to the list
        slides_shapes_info.append(slide_shape_info)

    scene = PresentationScene(slides_shapes_info, background_color)

    # Render the scene
    config.media_width = "100%"
    scene.render()

if __name__ == "__main__":
    main()
