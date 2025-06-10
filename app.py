from flask import Flask, request, send_file, jsonify
from PIL import Image
import io

app = Flask(__name__)

@app.route('/convert', methods=['POST'])
def convert_image():
    if 'image' not in request.files:
        return jsonify({'error': 'No image uploaded'}), 400

    image_file = request.files['image']
    target_format = request.form.get('format')

    if not target_format:
        return jsonify({'error': 'No target format specified'}), 400

    try:
        img = Image.open(image_file)

        # Convert to RGB if saving to JPEG
        if target_format.lower() in ['jpeg', 'jpg'] and img.mode in ("RGBA", "P"):
            img = img.convert("RGB")

        # Save to a BytesIO object
        img_bytes = io.BytesIO()
        img.save(img_bytes, format=target_format.upper())
        img_bytes.seek(0)

        # Set appropriate content type
        mime_type = f'image/{target_format.lower()}'
        return send_file(img_bytes, mimetype=mime_type, download_name=f'converted.{target_format.lower()}')

    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8080)

