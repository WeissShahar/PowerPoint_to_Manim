Work Hours





# Weekly Updates
#### Weeks 1-2: Introduction to Manim and Manim-Slides
During the initial phase, I focused on becoming proficient with the **Manim** and **Manim-Slides** libraries. This involved setting up the development environment, studying documentation, and creating a simple example project.

Key tasks completed:

* **Manim**: Explored Mobject creation, animations, and movement, and utilized the Manim grid for object positioning and transformations.
* **Manim-Slides**: Learned how to integrate and sequence multiple Manim scenes to create cohesive slideshows.

https://docs.manim.community/en/stable/tutorials/quickstart.html

https://manim-slides.eertmans.be/latest/reference/html.html

---

#### Week 3: Introduction to python-pptx
Following my exploration of Manim, the next objective was to extract data from a PowerPoint presentation. My first attempt was using the **python-pptx** library.

---

#### Week 4-5: Initial Shape Extraction and Manim Scene Creation
Using **python-pptx** and a basic PowerPoint presentation, I developed initial code to extract shapes from slides and recreate them as Mobjects in Manim.

Key details:

* The initial implementation focused on a single-slide presentation.
* Shape handling was limited to three types: Rectangle, Circle, and Triangle.
* Created the positioning transition, as PowerPoint works with (0,0) being top-left, and Manim works with (0,0) being the middle of the slide

During that time, I also improved it to work on a multiple slides presentation.

---

#### Week 6: Research on Color and Font Size Attributes

This phase presented significant challenges. Extracting attributes such as color, font size, and font from PowerPoint presentations using the python-pptx library proved difficult.

Key findings:

* For most shapes, the library did not return values for these attributes.
* In a few cases, I was able to retrieve some information by navigating the pptx hierarchy, but the results were inconsistent.

Following Week 6, each week included attempts to resolve this issue. These attributes are critical to replicating the exact presentation in Manim, and I have been continually researching and testing various solutions to overcome these challenges (More on that in the Achivments_And_issues.md file).

---

#### Week 7: Text Objects

Building on the effort to extract font size and font, I successfully completed the extraction of Text objects and their conversion into Manim Text objects.

Key tasks:

* Handled two types of text: text as a standalone object and text within a shape.
* Note: This was achieved without font size, font, or alignment, which will be addressed later in the conclusions.

---

#### Week 8: Animations
With a basic structure of shapes to mobjects, I turned my focus to adding animations. I primarily worked with position shifting and shape resizing animations, using the move_to command in Manim.

The animations were designed to show shapes moving between slides, demonstrating transitions. However, this approach proved problematic:

* Inconsistent shape IDs made it difficult to track objects across slides.
* Some objects were unintentionally animated, leading to unpredictable behavior.

As a result, I had to use an alternative approach: resetting the scene with each new slide, deleting all Mobjects and recreating the Mobjects for the following slide.

---

#### Week 9: Images and Lines 

Added support for Images and Line shapes.

---

#### Week 10: Shift to Code Generation Approach

Instead of directly generating a scene, I restructured the code to generate Manim code. This change allows for greater flexibility, enabling users to manually edit or run the scene after generation.

This approach offers more control and helps handle potential mismatches by allowing customization as needed.

---

#### Week 11: Tables

Added support for Tables.
However, as previously noted, the inability to extract font size from PowerPoint continues to cause a mismatch between the PowerPoint presentation and the generated Manim scene.

---

 #### Week 12: Arrows and Ellipses

 Added support for Arrows and Ellipses (replaced Oval-Circle).

 ---

 #### Weeks 13-14: XML

 During Weeks 13 and 14, I decided to dive deeper into XML parsing to extract more detailed information from PowerPoint files. My original goal was to figure out how to retrieve math equations, since the python-pptx library doesn't handle those well. But as I started working with the raw XML data, I realized there was more potential. 

This approach started to seem really promising, especially for things like font size, color, and alignment settings-all the details that python-pptx struggled with. XML has access to raw data directly from the PowerPoint files.

That being said, combining this XML-based method with my existing python-pptx workflow wasn't as easy. The XML structure is quite different from the more straightforward API that python-pptx provides, so trying to merge them was too challanging. It gave me more precision, but it also required a lot of manual handling, which made things more complicated.

Ultimately, working with XML showed me a much more accurate way to extract data. However, it also highlighted how difficult it would be to maintain the simplicity and flexibility of my original approach. Although I wasn't able to fully integrate XML parsing into my project this time, it clearly offers a promising path for future exploration. With further refinement, the XML method could significantly improve how  we replicate PowerPoint presentations in Manim, making it a valuable direction for future work.








