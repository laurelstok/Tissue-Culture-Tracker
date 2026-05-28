#!/usr/bin/env python3
"""
Confluency API Server — runs on the microscopy/analysis machine.
Exposes a simple HTTP API that TC Tracker calls to trigger image analysis.

Usage:
    pip install flask opencv-python scikit-image Pillow numpy pandas
    python3 confluency_api.py

Then set the URL (e.g. http://192.168.1.50:8082) in TC Tracker's settings.
"""

import os, re, json, traceback
from pathlib import Path
from datetime import datetime

try:
    from flask import Flask, request, jsonify
    from flask_cors import CORS
except ImportError:
    print("Missing flask. Run: pip install flask flask-cors")
    raise

try:
    import cv2
    import numpy as np
    from skimage import morphology
except ImportError:
    print("Missing cv2/skimage. Run: pip install opencv-python scikit-image numpy")
    raise

app = Flask(__name__)
CORS(app)

# ── Well ID parser ────────────────────────────────────────────────────────────
WELL_RE = re.compile(r'\b([A-P])([0-9]{1,2})\b', re.IGNORECASE)

def well_from_filename(name):
    """Extract well ID like A1, B12 from a filename."""
    m = WELL_RE.search(Path(name).stem)
    if m:
        return m.group(1).upper() + str(int(m.group(2)))
    return None

# ── Image analysis (adapted from ConfluencyAnalyzer) ─────────────────────────

def detect_dead_cells_raw(image, intensity_threshold=180):
    dead_cells_mask = np.zeros(image.shape, dtype=np.uint8)
    _, bright_mask = cv2.threshold(image, intensity_threshold, 255, cv2.THRESH_BINARY)
    clean_kernel = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (2, 2))
    bright_mask = cv2.morphologyEx(bright_mask, cv2.MORPH_OPEN, clean_kernel)
    contours, _ = cv2.findContours(bright_mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    for contour in contours:
        area = cv2.contourArea(contour)
        if not (20 < area < 3000): continue
        perimeter = cv2.arcLength(contour, True)
        if perimeter == 0: continue
        circularity = 4 * np.pi * area / (perimeter * perimeter)
        contour_mask = np.zeros(image.shape, dtype=np.uint8)
        cv2.fillPoly(contour_mask, [contour], 255)
        mean_intensity = np.mean(image[contour_mask > 0])
        if circularity > 0.55 and mean_intensity > intensity_threshold and 20 < area < 3000:
            cv2.fillPoly(dead_cells_mask, [contour], 255)
    return dead_cells_mask

def filter_dead_cells(image, mask, intensity_threshold=180):
    filtered_mask = mask.copy()
    contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    dead_cells_mask = np.zeros(mask.shape, dtype=np.uint8)
    for contour in contours:
        area = cv2.contourArea(contour)
        if area < 50: continue
        perimeter = cv2.arcLength(contour, True)
        if perimeter == 0: continue
        circularity = 4 * np.pi * area / (perimeter * perimeter)
        contour_mask = np.zeros(mask.shape, dtype=np.uint8)
        cv2.fillPoly(contour_mask, [contour], 255)
        mean_intensity = np.mean(image[contour_mask > 0])
        is_small_dead = circularity > 0.65 and mean_intensity > intensity_threshold and 50 < area < 2000
        is_bright_clump = circularity > 0.40 and mean_intensity > (intensity_threshold + 20) and 50 < area < 5000
        if is_small_dead or is_bright_clump:
            cv2.fillPoly(dead_cells_mask, [contour], 255)
            cv2.fillPoly(filtered_mask, [contour], 0)
    return filtered_mask, dead_cells_mask

def analyze_image(image_path, threshold_offset=-36, min_hole_size=200,
                  noise_kernel_size=5, dead_cell_threshold=180):
    """Analyze a single image, return confluency % and metadata."""
    img = cv2.imread(str(image_path), cv2.IMREAD_GRAYSCALE)
    if img is None:
        return None
    original_img = img.copy()
    early_dead_mask = detect_dead_cells_raw(original_img, dead_cell_threshold)
    kernel = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (noise_kernel_size, noise_kernel_size))
    img = cv2.morphologyEx(img, cv2.MORPH_OPEN, kernel)
    kernel = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (20, 20))
    img = cv2.morphologyEx(img, cv2.MORPH_TOPHAT, kernel)
    threshold_value, _ = cv2.threshold(img, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    adjusted_threshold = max(0, min(255, threshold_value + threshold_offset))
    _, binary = cv2.threshold(img, adjusted_threshold, 255, cv2.THRESH_BINARY)
    kernel = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (3, 3))
    binary = cv2.morphologyEx(binary, cv2.MORPH_CLOSE, kernel)
    binary_bool = binary > 0
    filled = morphology.remove_small_holes(binary_bool, area_threshold=min_hole_size)
    binary = (filled * 255).astype(np.uint8)
    binary, late_dead_mask = filter_dead_cells(original_img, binary, dead_cell_threshold)
    dead_cells_mask = cv2.bitwise_or(early_dead_mask, late_dead_mask)
    binary = cv2.bitwise_and(binary, cv2.bitwise_not(early_dead_mask))
    total_pixels = binary.size
    cell_pixels = int(np.sum(binary > 0))
    dead_cell_pixels = int(np.sum(dead_cells_mask > 0))
    confluency = min(100.0, round((cell_pixels / total_pixels) * 100, 2))
    return {
        'confluency': confluency,
        'cell_pixels': cell_pixels,
        'dead_cell_pixels': dead_cell_pixels,
        'total_pixels': total_pixels,
        'threshold_used': int(adjusted_threshold),
        'otsu_threshold': int(threshold_value),
        'mean_intensity': round(float(np.mean(original_img)), 2),
    }

# ── API endpoints ─────────────────────────────────────────────────────────────

@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok', 'service': 'confluency-api', 'version': '1.0'})

@app.route('/analyze', methods=['POST'])
def analyze():
    """
    POST /analyze
    Body: {
        "folder": "/path/to/images",
        "threshold_offset": -36,       # optional
        "min_hole_size": 200,           # optional
        "dead_cell_threshold": 180,     # optional
        "extensions": [".tif", ".png"]  # optional
    }
    Returns: {
        "results": [
            {"filename": "...", "well": "A1", "confluency": 72.3, ...},
            ...
        ],
        "summary": {"count": 6, "avg_confluency": 68.1, ...}
    }
    """
    try:
        body = request.get_json(force=True)
        folder = body.get('folder', '').strip()
        if not folder:
            return jsonify({'error': 'folder is required'}), 400

        folder_path = Path(folder)
        if not folder_path.exists():
            return jsonify({'error': f'Folder not found: {folder}'}), 404
        if not folder_path.is_dir():
            return jsonify({'error': f'Not a directory: {folder}'}), 400

        threshold_offset   = int(body.get('threshold_offset', -36))
        min_hole_size      = int(body.get('min_hole_size', 200))
        dead_cell_threshold = int(body.get('dead_cell_threshold', 180))
        extensions = [e.lower() for e in body.get('extensions', ['.tif', '.tiff', '.jpg', '.jpeg', '.png', '.bmp'])]

        image_files = [f for f in folder_path.iterdir()
                       if f.suffix.lower() in extensions and f.is_file()]

        if not image_files:
            return jsonify({'error': 'No image files found in folder', 'folder': str(folder_path)}), 404

        results = []
        for img_file in sorted(image_files):
            well = well_from_filename(img_file.name)
            measurements = analyze_image(
                img_file,
                threshold_offset=threshold_offset,
                min_hole_size=min_hole_size,
                dead_cell_threshold=dead_cell_threshold,
            )
            if measurements:
                results.append({
                    'filename': img_file.name,
                    'well': well,
                    'confluency': measurements['confluency'],
                    'dead_cell_pixels': measurements['dead_cell_pixels'],
                    'threshold_used': measurements['threshold_used'],
                    'mean_intensity': measurements['mean_intensity'],
                    'analyzed_at': datetime.now().isoformat(),
                })
            else:
                results.append({'filename': img_file.name, 'well': well, 'error': 'Analysis failed'})

        good = [r for r in results if 'confluency' in r]
        summary = {
            'count': len(results),
            'analyzed': len(good),
            'failed': len(results) - len(good),
            'avg_confluency': round(sum(r['confluency'] for r in good) / len(good), 1) if good else None,
            'min_confluency': round(min(r['confluency'] for r in good), 1) if good else None,
            'max_confluency': round(max(r['confluency'] for r in good), 1) if good else None,
            'folder': str(folder_path),
            'parameters': {
                'threshold_offset': threshold_offset,
                'min_hole_size': min_hole_size,
                'dead_cell_threshold': dead_cell_threshold,
            }
        }

        print(f'[analyze] {folder} → {len(good)}/{len(results)} images, avg confluency {summary["avg_confluency"]}%')
        return jsonify({'results': results, 'summary': summary})

    except Exception as e:
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser(description='Confluency Analysis API Server')
    parser.add_argument('--host', default='0.0.0.0', help='Host to listen on (default: 0.0.0.0)')
    parser.add_argument('--port', type=int, default=8082, help='Port (default: 8082)')
    args = parser.parse_args()
    print(f'Confluency API running at http://{args.host}:{args.port}')
    print('Endpoints: GET /health  POST /analyze')
    app.run(host=args.host, port=args.port, debug=False)