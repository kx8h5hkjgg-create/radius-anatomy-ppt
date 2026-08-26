# Python Script Documentation

## Overview

The `create_presentation.py` script generates a professional 30-slide PowerPoint presentation on radius bone anatomy using the `python-pptx` library.

## Prerequisites

- Python 3.6 or higher
- python-pptx library (version 0.6.21 or compatible)

## Installation

```bash
pip install python-pptx==0.6.21
```

## Usage

### Basic Usage
```bash
python create_presentation.py
```

### Output
Generates: `Radius_Anatomy_Seminar.pptx` (in current directory)

## Script Structure

### Color Scheme
```python
PRIMARY_COLOR = RGBColor(192, 0, 60)      # Pink/Maroon - titles
SECONDARY_COLOR = RGBColor(100, 149, 237) # Blue - alternatives
ACCENT_COLOR = RGBColor(220, 20, 60)      # Crimson - highlights
TEXT_COLOR = RGBColor(40, 40, 40)         # Dark gray - body text
LIGHT_BG = RGBColor(245, 245, 250)        # Off-white - background
```

### Main Functions

#### `create_presentation()`
- Creates presentation object
- Sets slide dimensions to 10" x 7.5"
- Returns: Presentation object

#### `add_title_slide(prs, title, subtitle)`
- Creates first/title slide
- Includes decorative top bar
- Displays main title and subtitle
- Adds footer with book reference
- Returns: Slide object

#### `add_content_slide(prs, title, content_list, slide_number=None)`
- Creates standard content slide
- Includes rounded rectangle title bar
- Displays bullet-pointed content
- Optional slide numbering
- Adds decorative bottom line
- Returns: Slide object

**Parameters:**
- `prs`: Presentation object
- `title`: Slide title string
- `content_list`: List of bullet points
- `slide_number`: Optional slide number to display

#### `add_two_column_slide(prs, title, left_title, left_content, right_title, right_content, slide_number=None)`
- Creates two-column comparison slide
- Each column has its own header and content
- Includes column separator line
- Returns: Slide object

**Parameters:**
- `prs`: Presentation object
- `title`: Main slide title
- `left_title`: Left column header
- `left_content`: List for left column
- `right_title`: Right column header
- `right_content`: List for right column
- `slide_number`: Optional slide number

#### `main()`
- Main execution function
- Creates all 30 slides
- Saves presentation file
- Prints status messages

## Slide Breakdown

| Slide # | Type | Title | Content |
|---------|------|-------|----------|
| 1 | Title | "THE RADIUS BONE" | Intro slide |
| 2 | Content | "Introduction to the Radius" | Basic facts |
| 3 | Content | "Basic Structure" | 3 main regions |
| 4 | Content | "Upper End (Proximal End)" | Head, neck, tuberosity |
| 5 | Content | "The Radial Head (Caput)" | Disc-like structure |
| 6 | Content | "Radial Head - Articular Surfaces" | Articulations |
| 7 | Content | "The Neck (Collum)" | Ligament support |
| 8 | Content | "Radial Tuberosity" | Biceps attachment |
| 9 | Content | "The Shaft (Diaphysis)" | Borders and surfaces |
| 10 | Content | "Anterior Border" | Muscle origins |
| 11 | Content | "Posterior Border" | Anatomical landmarks |
| 12 | Content | "Interosseous (Medial) Border" | Ulna connection |
| 13 | Content | "Anterior Surface of Shaft" | Muscle attachments |
| 14 | Content | "Posterior Surface of Shaft" | Flexor attachments |
| 15 | Content | "Lateral Surface of Shaft" | Nerve and muscle |
| 16 | Content | "Lower End (Distal End) - Overview" | 4 surfaces |
| 17 | Content | "Lateral Surface - Styloid Process" | Extends beyond ulna |
| 18 | Content | "Medial Surface - Ulnar Notch" | Inferior radio-ulnar joint |
| 19 | Content | "Anterior Surface of Lower End" | Radial artery |
| 20 | Content | "Posterior Surface - Dorsal Features" | Dorsal tubercle |
| 21 | Content | "Posterior Surface - Medial Grooves" | Extensor tendons |
| 22 | Content | "Posterior Surface - Additional Structures" | Retinaculum |
| 23 | Content | "Inferior Carpal Articular Surface" | Wrist joint |
| 24 | Content | "Articular Disc of Inferior Radio-Ulnar Joint" | Joint stability |
| 25 | Content | "Side Determination - How to Identify" | Bone identification |
| 26 | Content | "Key Articulations of Radius" | All joints |
| 27 | Two-Column | "Muscular Attachments" | Proximal & distal |
| 28 | Content | "Ligamentous Attachments" | All ligaments |
| 29 | Content | "Clinical Significance" | Fractures & conditions |
| 30 | Content | "Summary - Key Points to Remember" | Review |

## File Output

### Generated File
- **Filename**: `Radius_Anatomy_Seminar.pptx`
- **Format**: Microsoft PowerPoint 2007+ (.pptx)
- **Size**: Typically 200-500 KB
- **Compatibility**: MS Office, LibreOffice, Google Slides

### Console Output
```
🚀 Starting Radius Anatomy PowerPoint Generator...
📅 Generated on: 2026-08-26 05:25:20
Creating slide 1: Title Slide
Creating slide 2: Introduction
...
Creating slide 30: Summary

✅ Presentation created successfully!
📊 Total slides: 30
💾 File saved as: Radius_Anatomy_Seminar.pptx
📁 Location: /path/to/file

🎉 Ready for your anatomy seminar!
```

## Customization Guide

### Changing Colors

1. Locate color definitions at top of script:
```python
PRIMARY_COLOR = RGBColor(192, 0, 60)
```

2. Modify RGB values (0-255 each):
```python
PRIMARY_COLOR = RGBColor(0, 102, 204)  # Blue
ACCENT_COLOR = RGBColor(255, 153, 0)   # Orange
```

3. Regenerate presentation

### Modifying Content

Find the `add_content_slide()` call for desired slide:
```python
add_content_slide(prs, "Slide Title", [
    "• Original bullet point",
    "• Another original point"
], slide_num)
```

Modify text lists as needed.

### Changing Font Sizes

Locate font size definitions in each function:
```python
p.font.size = Pt(44)  # Title
p.font.size = Pt(20)  # Content
p.font.size = Pt(18)  # Slide number
```

Adjust Pt() values to desired size.

### Adding New Slides

Add after existing slides in `main()`:
```python
add_content_slide(prs, "New Slide Title", [
    "• Point 1",
    "• Point 2",
    "• Point 3"
], slide_num)
slide_num += 1
```

## Error Handling

The script includes try-except wrapper:
```python
if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()
```

## Debugging

### Enable Verbose Output
Add print statements in functions:
```python
def add_content_slide(prs, title, content_list, slide_number=None):
    print(f"Creating slide with title: {title}")
    # ... rest of function
```

### Check File Generation
```bash
ls -lh *.pptx  # On Linux/Mac
dir *.pptx     # On Windows
```

### Verify Installation
```bash
python -c "import pptx; print(pptx.__version__)"
```

## Performance Notes

- Script execution time: ~5-10 seconds
- File size: ~200-500 KB
- Memory usage: ~50-100 MB
- Slide rendering: Linear (1 slide per iteration)

## Dependencies

- **python-pptx**: PowerPoint generation library
  - Handles .pptx file creation
  - Manages slide layouts and shapes
  - Controls text formatting

## Code Structure

```
create_presentation.py
├── Imports
├── Color definitions
├── Function definitions
│   ├── create_presentation()
│   ├── add_title_slide()
│   ├── add_content_slide()
│   ├── add_two_column_slide()
│   └── main()
├── Main execution block
└── Error handling
```

## Best Practices

1. **Backup Original**: Keep original script before modifications
2. **Test Changes**: Regenerate after any modifications
3. **Version Control**: Use git for tracking changes
4. **Comments**: Add comments when customizing
5. **Testing**: Verify output in PowerPoint before use

## Limitations

- No built-in animations (add in PowerPoint)
- No images included (add in PowerPoint)
- No speaker notes (add in PowerPoint)
- Text-based content only
- Fixed slide dimensions

## Future Enhancements

Potential improvements:
- Add image embedding
- Include speaker notes
- Add animations programmatically
- Support multiple languages
- Interactive slide templates
- PDF export option

## Support

For issues or questions:
1. Check USAGE_GUIDE.md
2. Review README.md
3. Consult python-pptx documentation
4. Check Python error messages
5. Review script comments

## License

Free for educational use

## Version

- **Script Version**: 1.0
- **Created**: 2026-08-26
- **Python Requirement**: 3.6+
- **Library**: python-pptx 0.6.21+
