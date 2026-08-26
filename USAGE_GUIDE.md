# How to Use the Radius Anatomy PowerPoint Generator

## Quick Start

### 1. Install Dependencies
```bash
pip install -r requirements.txt
```

### 2. Generate the Presentation
```bash
python create_presentation.py
```

The script will create a file named `Radius_Anatomy_Seminar.pptx` in the same directory.

### 3. Open in PowerPoint
- Open the generated `.pptx` file with Microsoft PowerPoint, Google Slides, or LibreOffice Impress
- The presentation is ready to use for your seminar!

## Customization Guide

### Changing Colors

Edit the color definitions in `create_presentation.py`:

```python
PRIMARY_COLOR = RGBColor(192, 0, 60)  # Main title background
SECONDARY_COLOR = RGBColor(100, 149, 237)  # Alternative color
ACCENT_COLOR = RGBColor(220, 20, 60)  # Highlights
TEXT_COLOR = RGBColor(40, 40, 40)  # Body text
LIGHT_BG = RGBColor(245, 245, 250)  # Slide background
```

### Modifying Font Sizes

Adjust font sizes in individual slide functions:
- Title font: `Pt(44)` 
- Content font: `Pt(20)`
- Subtitle font: `Pt(32)`

### Editing Slide Content

Find the slide content lists and modify the text:

```python
add_content_slide(prs, "Slide Title", [
    "• Bullet point 1",
    "• Bullet point 2",
    "    - Sub-bullet",
], slide_num)
```

## Adding Animations

After opening the presentation in PowerPoint:

1. Click on a text box or shape
2. Go to **Animations** tab
3. Choose animation effect:
   - **Appear**: Simple fade-in
   - **Wipe**: Slides in from left
   - **Fade**: Gradual appearance
   - **Fly In**: Dynamic entrance

4. Set timing:
   - **On Click**: Appear when you click
   - **With Previous**: Appear with previous element
   - **After Previous**: Auto-play after previous

## Adding Images

### To add images to existing slides:

1. Open the presentation in PowerPoint
2. Go to the slide where you want to add an image
3. Click **Insert > Pictures**
4. Select your image file
5. Resize and position as needed

### Recommended images to add:
- Radius bone X-ray (anterior and posterior views)
- Anatomical diagram of radius structures
- Cross-sections of shaft at different levels
- Comparison with ulna bone

## Adding Speaker Notes

1. Open the presentation
2. Click **View > Notes**
3. Type detailed talking points in the notes section
4. Use for deeper explanations during the seminar

## Presentation Timing

**Recommended duration**: 45-60 minutes for complete coverage

- Slides 1-3: Introduction (5 minutes)
- Slides 4-8: Upper end (10 minutes)
- Slides 9-15: Shaft (10 minutes)
- Slides 16-24: Lower end (15 minutes)
- Slides 25-26: Identification & articulations (5 minutes)
- Slides 27-28: Attachments (5 minutes)
- Slides 29-30: Clinical & summary (5 minutes)

## Creating Handouts

### In PowerPoint:
1. Go to **File > Print**
2. Under **Settings**, select **Handouts**
3. Choose **6 slides per page** for optimal layout
4. Print or save as PDF

### Handout Content Suggestions:
- Anatomical terminology glossary
- Muscle attachment quick reference
- Clinical fracture types summary
- Wrist/hand movement guide

## Student Engagement Tips

1. **Interactive Questions**:
   - "Which bone is lateral - radius or ulna?"
   - "Where would you palpate for the radial pulse?"
   - "What fracture occurs from falling on outstretched hand?"

2. **Activities**:
   - Palpate own radius bone
   - Demonstrate pronation/supination
   - Identify anatomical landmarks on skeleton model

3. **Real-World Applications**:
   - Show X-rays of common fractures
   - Discuss wrist injuries in sports
   - Explain surgical approaches to radius

4. **Review Sessions**:
   - Use Slide 30 summary for review
   - Quiz on anatomical terms
   - Case study presentations

## Troubleshooting

### Script won't run
- Check Python version: `python --version` (should be 3.6+)
- Ensure python-pptx is installed: `pip install python-pptx`
- Run from correct directory

### File not created
- Check for error messages in console
- Ensure write permissions in directory
- Try running with administrator privileges

### Formatting issues
- Check if all fonts are available
- Verify color codes are correct RGB values
- Test with different PowerPoint versions

## Advanced Customization

### Adding new slide layouts:
```python
def add_special_slide(prs, title, custom_content):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    # Add your custom layout here
    return slide
```

### Modifying slide dimensions:
```python
prs.slide_width = Inches(10)  # Standard width
prs.slide_height = Inches(7.5)  # Standard height
```

### For more advanced features, refer to [python-pptx documentation](https://python-pptx.readthedocs.io/)

## Getting Help

- Review the README.md for general information
- Check inline comments in create_presentation.py
- Refer to python-pptx documentation for advanced modifications
- Test changes on a copy before final version

---

**Happy Teaching! 📚**
