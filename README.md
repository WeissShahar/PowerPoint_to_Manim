# PowerPoint To Manim Project

This repository contains my Computer Graphics project, which aims to generate Manim code from PowerPoint presentations. Using python-pptx to extract slide data, alongside Manim and Manim-Slides for rendering, the project seeks to recreate PowerPoint slides as accurately as possible, while addressing the limitations of each tool.


## How to Run

### Prerequisites:
- Python 3.x
- Manim (Installation guide: [Manim Docs](https://docs.manim.community/en/stable/installation.html))
- python-pptx 
- manim-slides

### Setup:
1. Clone the repository:
   ```bash
   git clone https://github.com/yourusername/PowerPoint-to-Manim.git
   cd PowerPoint-to-Manim
2. Install the required packages:

   ```bash 
   pip install -r requirements.txt
### Running the script:

You can run the script with optional arguments to choose which stages to execute:
```bash
    python main.py /YourPresentation.pptx
```

* No arguments will generate the Manim code.
* Use --render to render the Manim scene.
* Use --convert to convert the rendered scene to PPTX and HTML and open it in a browser.

### Important Note:
- You must run the rendering stage before converting. The conversion will not work without first rendering the Manim code.

## Post Execution
After running the script, a new Python file named **generated_manim_code_for_{PresentationName}.py** will be created. The generated file will contain code that uses the Manim library to render the PowerPoint slides as scenes.

The file will contain something like this:
```bash
class GeneratedPresentation(Slide):
    def construct(self):
        self.camera.background_color = WHITE
        config.frame_width = 13.333333333333334
        config.frame_height = 7.5
    
        # Slide 1
        self.clear()
        mobject = Text('''Path Planning''', font_size=24, color=BLACK)
        mobject.move_to([0.0, 0.4130435258092737, 0])
        self.add(mobject)
        mobject.move_to([0.0, 0.4130435258092737, 0])
        self.add(mobject)
        self.wait(1)
        self.next_slide()

        # Slide 2
        self.clear()
        mobject = Text('''Problem Statement – 2D environment''', font_size=24, color=BLACK)
        mobject.move_to([0.0, 2.6258677821522314, 0])
        self.add(mobject)
        mobject.move_to([0.0, 2.6258677821522314, 0])
        self.add(mobject)
        mobject = ImageMobject('extracted_images\slide_1_image_11.png')
        mobject.width, mobject.height = (11.179761592300963, 3.912047244094488)
        mobject.move_to([0.0, -0.31637357830271196, 0])
        self.add(mobject)
        self.wait(1)
        self.next_slide()
```

This is the generated Manim code. you can modify it (change data, positioning, colors and more), or you can run it as it is.

In addition, a directory named **extracted_images** will be created, where each image from the PowerPoint will be saved. The images will be named in the format slide_{i}_image_{j}.png, where i is the slide number, and j is the image index within that slide.

### --render
If you used the --render option, your Manim scene will also be created. The rendered video of your slides will be located in the .../media/videos directory.

If you haven't rendered it yet, you can do so by running the following command:

```bash 
manim ql {script_name.py} GeneratedPresentation
```

### --convert
If you used the --convert option, the script will convert your rendered Manim scene into a PPTX file and an HTML file and automatically open it in your browser. 

The generated PPTX file HTML file will be named manim_presentation.XXX and will be located in the same directory where you ran the script.

If you haven't converted it yet but have already rendered the scene, you can convert it by running the following command:

``` bash
manim-slides convert GeneratedPresentation manim_presentation.html
```

``` bash
manim-slides convert --to pptx GeneratedPresentation manim_presentation.pptx
```

## Customization Options:

Currently the project supports:
* Rectangle
* Oval
* Line
* Arrow
* Text Box
* Image
* Table

In order to add a new shape, follow these steps:

1. Handle the Shape in the extract_shapes_from_slide Function:
* Modify the extract_shapes_from_slide function to identify the new shape and extract the relevant properties, such as dimensions, position, or other specific attributes.

Example for adding a Star:
```bash
if shape.auto_shape_type == MSO_AUTO_SHAPE_TYPE.STAR:
    shape_info['type'] = 'star'
    ...extract info...
```

2. Generate Manim Code for the Shape:
*  Add a corresponding code generation block in the generate_manim_code function, similar to how other shapes are handled.

```bash
elif shape_info['type'] == 'star':
    slide_code += f"        mobject = Star(...)\n"
```

## Results:

For more details on the challenges and issues encountered during the process, see the [Achievements_And_Issues](https://github.com/WeissShahar/PowerPoint_to_Manim/blob/master/Achievements_And_Issues.md) section.

Below are two examples comparing PowerPoint slides with their corresponding Manim-generated slides:

1. [Example 1](https://draftable.com/compare/BkarpogpdKxm)
2. [Example 2](https://draftable.com/compare/LPBhImDhwuka)


In these comparisons, the left side shows the original PowerPoint slide, and the right side displays the slide generated by Manim.

Important Note:
Converting a Manim-generated slide back to a PowerPoint format will result in the loss of any animations created in Manim. However, since our project does not include animations, no content is lost in these examples.


## XML Extraction
Check the [XML_Testing](https://github.com/WeissShahar/PowerPoint_to_Manim/blob/master/xml_testing.py) for the code.


While python-pptx offered a basic way to extract data from PowerPoint slides, it struggled to retrieve key details like font sizes, colors, and alignment. XML parsing, on the other hand, taps directly into the raw data within PowerPoint files, offering far more precision. This approach allows access to all attributes and slide elements that python-pptx often missed, making it possible to more faithfully replicate slides in Manim.

With XML parsing, I was able to extract more accurate font sizes, detailed color values, and shape properties, leading to significantly improved scene recreation in Manim. However, fully integrating XML parsing into the project wasn't achiveable within the current scope, meaning this version remains in a BETA state. Although still in testing, it offers a promising path forward for future developments.

Below are two examples that demonstrate the results using XML parsing:

**PPTX** (left) | **Manim** (right)
![XML_1](https://github.com/user-attachments/assets/a0df064d-b7ab-4e64-91b1-fe9041a3560a)

---

**PPTX** (left) | **Manim** (right)
![XML_2](https://github.com/user-attachments/assets/db77d6dc-cd21-4fdf-a43c-0a7b9d2a9ca3)

As we can see, the color match-making is possible, but needs a little bit more handling


## Conclusion:

This project set out to bridge the gap between PowerPoint presentations and Manim animations, aiming to generate accurate Manim scenes from slide data. While python-pptx provided a solid foundation, it revealed limitations in extracting certain key details like font sizes, colors, and alignments. However, this challenge opened the door to discovering a more precise method—XML parsing.

Through XML parsing, I was able to directly access the raw data within PowerPoint files, offering much greater accuracy and control over the content. This approach significantly improves the fidelity of recreating slides in Manim and sets a promising path forward.

Looking ahead, focusing on XML extraction and implementing animation transitions between PowerPoint and Manim will bring the project even closer to its full potential. Overall, the project has laid a strong foundation for further development, with exciting possibilities for future enhancements
