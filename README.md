# PPT_PROJ

from pathlib import Path
from PIL import Image, ImageFilter
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor


# ============================================================
# 1. MAIN FOLDER
# ============================================================

main_folder = Path.cwd() / "CFD_PPT_PROJ"

if not main_folder.exists():
    raise FileNotFoundError(
        f"Main folder not found:\n{main_folder}"
    )

print("Main folder:")
print(main_folder)


# ============================================================
# 2. SOURCE IMAGE FOLDERS
# ============================================================

folders = {
    "velocity_case1":
        main_folder / "Case1" / "cf Rig" / "Velocity side view",

    "velocity_case2":
        main_folder / "Case2" / "cf Rig" / "Velocity side view",

    "pressure_case1":
        main_folder / "Case1" / "cf Rig" / "Pressure Side View",

    "pressure_case2":
        main_folder / "Case2" / "cf Rig" / "Pressure Side View",
}


# ============================================================
# 3. CROPPED IMAGE FOLDER
# ============================================================

cropped_root = main_folder / "Cropped_Images"
cropped_root.mkdir(exist_ok=True)


# ============================================================
# 4. GET IMAGES
# ============================================================

def get_images(folder):

    images = [
        p for p in folder.iterdir()
        if p.is_file()
        and p.suffix.lower() in [".jpg", ".jpeg", ".png"]
    ]

    def image_number(path):
        digits = ''.join(
            filter(str.isdigit, path.stem)
        )
        return int(digits) if digits else 9999

    images.sort(key=image_number)

    if len(images) < 5:
        raise ValueError(
            f"\n{folder}\n"
            f"contains only {len(images)} images.\n"
            f"5 images are required."
        )

    return images[:5]


# ============================================================
# 5. AUTOMATIC IMAGE CROP
# ============================================================

def auto_crop_image(
    input_path,
    output_path,
    margin=15,
    threshold=245
):

    img = Image.open(input_path).convert("RGB")

    gray = img.convert("L")

    # Reduce tiny noise
    gray = gray.filter(
        ImageFilter.GaussianBlur(radius=2)
    )

    # Detect non-white area
    mask = gray.point(
        lambda p: 255 if p < threshold else 0
    )

    bbox = mask.getbbox()

    if bbox is None:

        img.save(
            output_path,
            quality=95
        )

        return

    left, top, right, bottom = bbox

    # Safety margin
    left = max(0, left - margin)
    top = max(0, top - margin)
    right = min(img.width, right + margin)
    bottom = min(img.height, bottom + margin)

    cropped = img.crop(
        (
            left,
            top,
            right,
            bottom
        )
    )

    cropped.save(
        output_path,
        quality=95
    )


# ============================================================
# 6. LOAD AND CROP ALL REQUIRED IMAGES
# ============================================================

cropped_images = {}

for name, folder in folders.items():

    if not folder.exists():
        raise FileNotFoundError(
            f"Folder not found:\n{folder}"
        )

    original_images = get_images(folder)

    output_folder = cropped_root / name
    output_folder.mkdir(exist_ok=True)

    cropped_images[name] = []

    for image_path in original_images:

        output_path = (
            output_folder / image_path.name
        )

        auto_crop_image(
            image_path,
            output_path
        )

        cropped_images[name].append(
            output_path
        )

    print(
        f"{name}: {len(cropped_images[name])} images"
    )


# ============================================================
# 7. CREATE POWERPOINT
# ============================================================

prs = Presentation()

# 16:9
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

slide = prs.slides.add_slide(
    prs.slide_layouts[6]
)


# ============================================================
# 8. COLORS
# ============================================================

BLACK = RGBColor(0, 0, 0)
WHITE = RGBColor(255, 255, 255)

DARK_BLUE = RGBColor(20, 70, 120)
LIGHT_BLUE = RGBColor(220, 235, 250)

GRID_BLUE = RGBColor(80, 130, 180)

LIGHT_GRAY = RGBColor(245, 245, 245)


# ============================================================
# 9. SLIDE TITLE
# ============================================================

title = slide.shapes.add_shape(
    MSO_SHAPE.RECTANGLE,
    Inches(0.05),
    Inches(0.05),
    Inches(13.23),
    Inches(0.45)
)

title.fill.solid()
title.fill.fore_color.rgb = DARK_BLUE

title.line.fill.background()

tf = title.text_frame
tf.clear()
tf.vertical_anchor = MSO_ANCHOR.MIDDLE

p = tf.paragraphs[0]
p.text = "CFD COMPARISON"
p.alignment = PP_ALIGN.CENTER

run = p.runs[0]
run.font.size = Pt(20)
run.font.bold = True
run.font.color.rgb = WHITE


# ============================================================
# 10. LAYOUT DIMENSIONS
# ============================================================

# Left section-name column
section_x = 0.05
section_width = 1.55

# Case column
case_x = 1.62
case_width = 0.95

# Image columns
image_start_x = 2.62
image_width = 2.10
image_height = 1.25

column_gap = 0.05


# ============================================================
# 11. LIFT HEADER
# ============================================================

header_y = 0.60
header_height = 0.42


def add_header_cell(
    text,
    x,
    width
):

    shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(x),
        Inches(header_y),
        Inches(width),
        Inches(header_height)
    )

    shape.fill.solid()
    shape.fill.fore_color.rgb = LIGHT_BLUE

    shape.line.color.rgb = GRID_BLUE
    shape.line.width = Pt(1.2)

    tf = shape.text_frame
    tf.clear()
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE

    p = tf.paragraphs[0]
    p.text = text
    p.alignment = PP_ALIGN.CENTER

    run = p.runs[0]
    run.font.size = Pt(13)
    run.font.bold = True
    run.font.color.rgb = BLACK


# Lift
add_header_cell(
    "Lift",
    section_x,
    section_width + case_width
)


# Lift values
lift_values = [
    "2 mm",
    "4 mm",
    "6 mm",
    "8 mm",
    "10 mm"
]

for i, lift in enumerate(lift_values):

    x = (
        image_start_x
        + i * (image_width + column_gap)
    )

    add_header_cell(
        lift,
        x,
        image_width
    )


# ============================================================
# 12. IMAGE FUNCTION
# ============================================================

def add_image_contained(
    image_path,
    x,
    y,
    width,
    height
):

    img = Image.open(image_path)

    img_width, img_height = img.size

    image_ratio = (
        img_width / img_height
    )

    box_ratio = (
        width / height
    )

    if image_ratio > box_ratio:

        final_width = width
        final_height = width / image_ratio

    else:

        final_height = height
        final_width = height * image_ratio

    final_x = (
        x + (width - final_width) / 2
    )

    final_y = (
        y + (height - final_height) / 2
    )

    slide.shapes.add_picture(
        str(image_path),
        Inches(final_x),
        Inches(final_y),
        width=Inches(final_width),
        height=Inches(final_height)
    )


# ============================================================
# 13. CASE CELL
# ============================================================

def add_case_cell(
    case_name,
    y
):

    shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(case_x),
        Inches(y),
        Inches(case_width),
        Inches(image_height)
    )

    shape.fill.solid()
    shape.fill.fore_color.rgb = LIGHT_BLUE

    shape.line.color.rgb = GRID_BLUE
    shape.line.width = Pt(1.2)

    tf = shape.text_frame
    tf.clear()
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE

    p = tf.paragraphs[0]
    p.text = case_name
    p.alignment = PP_ALIGN.CENTER

    run = p.runs[0]
    run.font.size = Pt(13)
    run.font.bold = True
    run.font.color.rgb = BLACK


# ============================================================
# 14. IMAGE ROW
# ============================================================

def add_image_row(
    image_list,
    case_name,
    y
):

    # Case label
    add_case_cell(
        case_name,
        y
    )

    # Five CFD images
    for i, image_path in enumerate(image_list):

        x = (
            image_start_x
            + i * (image_width + column_gap)
        )

        # Grid box
        cell = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Inches(x),
            Inches(y),
            Inches(image_width),
            Inches(image_height)
        )

        cell.fill.solid()
        cell.fill.fore_color.rgb = WHITE

        cell.line.color.rgb = GRID_BLUE
        cell.line.width = Pt(1.2)

        # Image
        add_image_contained(
            image_path,
            x,
            y,
            image_width,
            image_height
        )


# ============================================================
# 15. SECTION LABEL
# ============================================================

def add_section_label(
    text,
    y,
    total_height
):

    # Outer cell
    outer = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(section_x),
        Inches(y),
        Inches(section_width),
        Inches(total_height)
    )

    outer.fill.solid()
    outer.fill.fore_color.rgb = WHITE

    outer.line.color.rgb = GRID_BLUE
    outer.line.width = Pt(1.2)


    # Inner highlighted label
    inner = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(section_x + 0.15),
        Inches(
            y + total_height / 2 - 0.40
        ),
        Inches(section_width - 0.30),
        Inches(0.80)
    )

    inner.fill.solid()
    inner.fill.fore_color.rgb = DARK_BLUE

    inner.line.fill.background()

    tf = inner.text_frame
    tf.clear()
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE

    p = tf.paragraphs[0]

    # Split into two lines
    if "VELOCITY" in text:
        p.text = "VELOCITY\nSIDE VIEW"
    else:
        p.text = "PRESSURE\nSIDE VIEW"

    p.alignment = PP_ALIGN.CENTER

    run = p.runs[0]
    run.font.size = Pt(11)
    run.font.bold = True
    run.font.color.rgb = WHITE


# ============================================================
# 16. VELOCITY SECTION
# ============================================================

velocity_y = 1.08

add_section_label(
    "VELOCITY SIDE VIEW",
    velocity_y,
    image_height * 2
)

add_image_row(
    cropped_images["velocity_case1"],
    "Case 1",
    velocity_y
)

add_image_row(
    cropped_images["velocity_case2"],
    "Case 2",
    velocity_y + image_height
)


# ============================================================
# 17. PRESSURE SECTION
# ============================================================

pressure_y = 3.78

add_section_label(
    "PRESSURE SIDE VIEW",
    pressure_y,
    image_height * 2
)

add_image_row(
    cropped_images["pressure_case1"],
    "Case 1",
    pressure_y
)

add_image_row(
    cropped_images["pressure_case2"],
    "Case 2",
    pressure_y + image_height
)


# ============================================================
# 18. SAVE POWERPOINT
# ============================================================

output_file = (
    main_folder
    / "CFD_COMPARISON_FINAL.pptx"
)

prs.save(output_file)


# ============================================================
# 19. DONE
# ============================================================

print()
print("==============================================")
print("POWERPOINT CREATED SUCCESSFULLY")
print("==============================================")
print()
print("PowerPoint:")
print(output_file)
print()
print("Cropped images:")
print(cropped_root)
