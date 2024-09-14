from pptx import Presentation
import xml.etree.ElementTree as ET
from lxml import etree

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

    # Parse the slide XML
    root = etree.fromstring(slide_xml)

    # Find all shapes
    shapes = root.findall('.//p:sp', namespaces=ns)
    for shape in shapes:
        info = {}

        # Extract shape type
        shape_type = shape.find('.//p:nvSpPr/p:cNvPr', namespaces=ns)
        if shape_type is not None:
            shape_name = shape_type.get('name')
            if shape_name and 'Oval' in shape_name:
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

        # Extract color
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
                    lum_mod=int(scheme_clr.find('a:lumMod', namespaces=ns).get('val')) if scheme_clr.find('a:lumMod', namespaces=ns) is not None else None,
                    shade_mod=int(scheme_clr.find('a:shade', namespaces=ns).get('val')) if scheme_clr.find('a:shade', namespaces=ns) is not None else None,
                    theme_colors=theme_colors
                )

        # Extract text content
        text_elems = shape.findall('.//a:t', namespaces=ns)
        if text_elems:
            text_content = '\n'.join([elem.text for elem in text_elems if elem.text is not None])
            info['text'] = text_content

        # Extract font size and text color
        rPr = shape.find('.//a:rPr', namespaces=ns)
        if rPr is not None:
            info['font_size'] = int(rPr.get('sz')) / 100  # Font size is in 100ths of a point
            text_fill = rPr.find('.//a:solidFill/a:schemeClr', namespaces=ns)
            if text_fill is not None:
                text_color_val = text_fill.get('val')
                # Map to actual color using master_slide_clr_mapping
                actual_text_color_val = master_slide_clr_mapping.get(text_color_val, text_color_val)
                # Get the actual text color value from theme_colors
                info['text_color'] = {
                    'val': get_adjusted_color(
                        actual_text_color_val,
                        lum_mod=int(text_fill.find('a:lumMod', namespaces=ns).get('val')) if text_fill.find('a:lumMod', namespaces=ns) is not None else None,
                        shade_mod=int(text_fill.find('a:shade', namespaces=ns).get('val')) if text_fill.find('a:shade', namespaces=ns) is not None else None,
                        theme_colors=theme_colors
                    )
                }

        if info:
            shape_info.append(info)

    return shape_info

EMU_TO_POINTS = 1 / 12700

def convert_position(shape, slide_width_in_points, slide_height_in_points):
    """Convert PowerPoint coordinates to Manim coordinates."""

    # Convert from EMU to points
    x_in_points = shape['position']['x'] * EMU_TO_POINTS
    y_in_points = shape['position']['y'] * EMU_TO_POINTS

    # Convert to Manim coordinates
    manim_x = x_in_points / 72 - slide_width_in_points / 2 + shape['size']['cx'] * EMU_TO_POINTS / 2 / 72
    manim_y = slide_height_in_points / 2 - (y_in_points / 72 + shape['size']['cy'] * EMU_TO_POINTS / 2 / 72)

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

def main():
    # Load the presentation
    presentation = Presentation('presentations/test.pptx')

    # Get the theme part from the related parts of the presentation
    theme_part = None
    for rel in presentation.part.rels.values():
        if "theme" in rel.target_ref:
            theme_part = rel.target_part
            break

    # Get the theme XML (if any)
    theme_xml = theme_part._blob.decode() if theme_part else None
    slide = presentation.slides[1]
    slide_xml = slide._element.xml

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
    shape_info = extract_shape_info(slide_xml, master_slide_clr_mapping, theme_colors)

    # Print the final color mapping
    print("Final Color Mapping:")
    for master_color, value in final_color_mapping.items():
        print(f"{master_color}: {value}")

    # Print the shape info
    print("Shape Info:")
    print(shape_info)

if __name__ == "__main__":
    main()
