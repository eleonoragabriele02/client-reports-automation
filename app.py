"""
Client Reports Automation - Microservice
========================================

Flask microservice that generates .pptx reports from templates by:
1. Replacing {{placeholder}} text tags with actual data
2. Replacing images on specified slides with chart images from URLs

Endpoints:
- GET  /          : health check
- POST /generate  : generate a report from JSON body

Deployed on Render.com free tier.
"""
from flask import Flask, request, send_file, jsonify
from pptx import Presentation
import requests
import io
import os
import traceback

app = Flask(__name__)

TEMPLATES_DIR = os.path.dirname(os.path.abspath(__file__))
REQUEST_TIMEOUT = 30  # seconds for image download


# ==========================================
# TEXT REPLACEMENT
# ==========================================

def replace_placeholders_in_paragraph(paragraph, placeholders):
    """
    Replace {{placeholder}} tags in a paragraph, handling text runs
    that may be fragmented across multiple <a:r> elements in the XML.
    """
    # Concatenate all runs to get the full paragraph text
    full_text = ''.join(run.text for run in paragraph.runs)
    if '{{' not in full_text:
        return False

    new_text = full_text
    for key, value in placeholders.items():
        marker = '{{' + key + '}}'
        if marker in new_text:
            replacement = str(value) if value is not None else ''
            new_text = new_text.replace(marker, replacement)

    if new_text == full_text:
        return False

    # Consolidate: put all new text in first run, clear others
    if paragraph.runs:
        paragraph.runs[0].text = new_text
        for run in paragraph.runs[1:]:
            run.text = ''

    return True


def replace_in_shape(shape, placeholders):
    """Recursively process a shape for text replacement."""
    # Regular text frames
    if shape.has_text_frame:
        for para in shape.text_frame.paragraphs:
            replace_placeholders_in_paragraph(para, placeholders)

    # Table cells
    if shape.has_table:
        for row in shape.table.rows:
            for cell in row.cells:
                for para in cell.text_frame.paragraphs:
                    replace_placeholders_in_paragraph(para, placeholders)

    # Groups - recurse into sub-shapes
    try:
        if shape.shape_type == 6:  # GROUP
            for sub_shape in shape.shapes:
                replace_in_shape(sub_shape, placeholders)
    except Exception:
        pass


# ==========================================
# IMAGE REPLACEMENT
# ==========================================

def replace_image_on_slide(slide, image_bytes, image_index=0):
    """
    Replace the Nth largest picture on a slide by removing it and
    adding a new picture at the same position/size.
    """
    pictures = [s for s in slide.shapes if s.shape_type == 13]  # PICTURE = 13
    if not pictures:
        return False

    # Sort by area descending (largest first — the chart image is typically the biggest)
    pictures.sort(key=lambda p: (p.width or 0) * (p.height or 0), reverse=True)

    if image_index >= len(pictures):
        return False

    target = pictures[image_index]
    left, top = target.left, target.top
    width, height = target.width, target.height

    # Remove old picture from slide XML
    sp = target._element
    sp.getparent().remove(sp)

    # Add new picture at same position/size
    slide.shapes.add_picture(io.BytesIO(image_bytes), left, top, width, height)
    return True


# ==========================================
# ENDPOINTS
# ==========================================

@app.route('/', methods=['GET'])
def health():
    """Health check endpoint. Also lists available templates."""
    templates = [f for f in os.listdir(TEMPLATES_DIR) if f.endswith('.pptx')]
    return jsonify({
        'status': 'ok',
        'service': 'client-reports-automation',
        'templates_dir': TEMPLATES_DIR,
        'available_templates': templates
    })


@app.route('/generate', methods=['POST'])
def generate_report():
    """
    Generate a pptx report from template + data.

    Request body (JSON):
    {
        "template": "Malne-Template.pptx",
        "placeholders": {
            "tra_web_curr": "27.729",
            "comentario_trafico": "...",
            ...
        },
        "images": {
            "5": "https://quickchart.io/chart?...",
            "6": "https://quickchart.io/chart?..."
        }
    }
    """
    try:
        data = request.get_json(force=True, silent=False)
        if not data:
            return jsonify({'error': 'Empty or invalid JSON body'}), 400

        template_name = data.get('template', 'Malne-Template.pptx')
        placeholders = data.get('placeholders', {})
        images = data.get('images', {})

        # Validate template exists
        template_path = os.path.join(TEMPLATES_DIR, template_name)
        if not os.path.exists(template_path):
            available = [f for f in os.listdir(TEMPLATES_DIR) if f.endswith('.pptx')]
            return jsonify({
                'error': f'Template not found: {template_name}',
                'available': available
            }), 404

        # Load template
        prs = Presentation(template_path)

        # Replace text placeholders across all slides
        for slide in prs.slides:
            for shape in slide.shapes:
                replace_in_shape(shape, placeholders)

        # Replace images on specified slides
        images_replaced = 0
        for slide_num_str, image_url in images.items():
            try:
                slide_num = int(slide_num_str)
                slide_idx = slide_num - 1
                if slide_idx < 0 or slide_idx >= len(prs.slides):
                    print(f'[WARN] Slide {slide_num} out of range')
                    continue

                print(f'[INFO] Downloading image for slide {slide_num}...')
                resp = requests.get(image_url, timeout=REQUEST_TIMEOUT)
                if resp.status_code != 200:
                    print(f'[WARN] Image download failed slide {slide_num}: HTTP {resp.status_code}')
                    continue

                slide = prs.slides[slide_idx]
                if replace_image_on_slide(slide, resp.content):
                    images_replaced += 1
                    print(f'[OK] Image replaced on slide {slide_num}')
                else:
                    print(f'[WARN] No picture found on slide {slide_num} to replace')
            except Exception as e:
                print(f'[ERROR] Image replacement slide {slide_num_str}: {e}')
                continue

        # Save to memory buffer
        output = io.BytesIO()
        prs.save(output)
        output.seek(0)

        # Build filename
        template_base = os.path.splitext(template_name)[0]
        filename = f'{template_base}-generated.pptx'

        print(f'[SUCCESS] Report generated. Images replaced: {images_replaced}')

        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation',
            as_attachment=True,
            download_name=filename
        )

    except Exception as e:
        traceback.print_exc()
        return jsonify({
            'error': str(e),
            'traceback': traceback.format_exc().split('\n')[-10:]
        }), 500


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
