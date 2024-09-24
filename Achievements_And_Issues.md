## Achievements and Issues

As already mentioned, throughout this project, using python-pptx presented significant challenges when extracting key information such as font styles, font sizes, background colors, shape colors, and text alignment. These limitations often resulted in major inaccuracies when attempting to replicate PowerPoint slides within Manim.

In this section, we'll mention some achivments made and explore the most notable issues encountered.

### Achievements:
**Accurate Positioning of Shapes:**

The project successfully extracted shape positions and translated them into corresponding locations within Manim's coordinate system. This allowed for consistent layout replication.

**Image Handling:**

Images embedded in PowerPoint slides were successfully extracted and rendered within the Manim scene, retaining their correct positions and proportions.

**Data Extraction:**

The extraction of most non-styled data, such as shape types, text content, and basic attributes (e.g., size, position), was achieved, ensuring a solid framework for continued improvements in detailed styling.

### Issues:

### Font, Font Size, and Newline:
**PPTX** (left) | **Manim** (right)

![Issues_1](https://github.com/user-attachments/assets/7ba5e677-bb81-43c6-ba3b-ee54f43d1565)

---

**Font and Font Size**: Inability to extract accurate font and size information led to mismatches, defaulting to placeholder values.

**Possible Solution**: Use XML parsing to retrieve precise font details.

---

**Newline Issue**: Textboxes didn't recognize newlines, causing all text to appear on a single line. This happened because newlines were determined by the textbox dimensions.

**Possible Solution**: Calculate the textbox dimensions and check if words exceed the width to insert newlines accordingly.

---

### Colors and Text Alignment:

**PPTX** (left) | **Manim** (right)
![Issues_2](https://github.com/user-attachments/assets/6473bbc7-667b-4828-9bf4-0eae38f17df4)

---

**Colors**: Failure to extract shape and background colors resulted in defaulting to white backgrounds and black text.

**Possible Solution**: Use XML parsing to retrieve precise color information.

---

**Alignment Issues**: Text in full-screen text boxes was centered by default due to missing alignment data.

**Possible Solution**: Use XML parsing to extract exact text alignment.

---

**Additionally**, Manim centers text within shapes by default, potentially causing overlap.

**Possible Solution**: Manually adjust text position after generating the Manim code.

---

### Tables:

**PPTX** (left) | **Manim** (right)
![Issues_3](https://github.com/user-attachments/assets/dcab59e0-5c8f-4fdb-8d34-a6c697aa24bb)

---

**Tables**: Inability to extract cell sizes, and Manim's use of scale to change table size.

**Possible Solution:** Extract cells size with XML, and manually adjust tables size and location as needed after generating the Manim code.

---

### Math Equation and Lines:

**PPTX** (left) | **Manim** (right)
![Issues_4](https://github.com/user-attachments/assets/203e84b1-c6c8-469f-af13-b8b535f02641)

---

**Math Equations:** 
Math Equations not being supported in python-pptx

**Possible solution**: Use XML parsing to extract Math Equations.
Note: Converting it to Manim might still be an issue.

---

**Lines**:
Lines accuracy is not consistent through out the slides.

**Possible Solution:** Find a better generalizing technique for Lines and Arrows

---

### General Issues:

1. **Shape Disparity:**

    Issue: PPTX files often contain a wider variety of shapes and object types than what is supported in Manim. This includes specialized shapes like SmartArt, charts, and embedded objects.

2. **Shape Addition:**

    Issue: The need to manually add support for every shape type that may appear in a PowerPoint presentation, which can be time-consuming and may lead to inconsistencies.

3. **Animation and Transition Support:**

    Issue: Manim may not support all the animation and transition effects available in PowerPoint, leading to a loss of visual dynamics.
    **Note:** The original goal of this project included support for animations, but this feature was not implemented. As a result, a basic method for handling animation transitions is still needed.

4. **Embedded Media:**
    Issue:  PPTX files may contain embedded audio or video, which cannot be directly represented in Manim.


