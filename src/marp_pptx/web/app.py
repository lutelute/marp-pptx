"""Flask web UI for marp-pptx."""
from __future__ import annotations

import hashlib
import shutil
import subprocess
import tempfile
import uuid
from pathlib import Path

from flask import Flask, request, send_file, render_template, jsonify


# Cache for rendered slide thumbnails (keyed by MD content hash + settings)
_PREVIEW_CACHE_DIR = Path(tempfile.gettempdir()) / "marp_pptx_previews"
_PREVIEW_CACHE_DIR.mkdir(exist_ok=True)

_SOFFICE = shutil.which("soffice") or shutil.which("libreoffice")
_PDFTOPPM = shutil.which("pdftoppm")


def _render_pptx_to_pngs(pptx_path: Path, out_dir: Path, dpi: int = 100) -> list[Path]:
    """Convert a PPTX file to per-slide PNG thumbnails using soffice + pdftoppm.

    Returns the sorted list of PNG paths. Returns [] if tools unavailable
    or conversion fails.
    """
    if _SOFFICE is None or _PDFTOPPM is None:
        return []
    out_dir.mkdir(parents=True, exist_ok=True)
    try:
        # Step 1: PPTX → PDF via LibreOffice
        subprocess.run(
            [_SOFFICE, "--headless", "--convert-to", "pdf",
             "--outdir", str(out_dir), str(pptx_path)],
            check=True, capture_output=True, timeout=60,
        )
        pdf = out_dir / (pptx_path.stem + ".pdf")
        if not pdf.exists():
            return []
        # Step 2: PDF → PNG per page via pdftoppm
        subprocess.run(
            [_PDFTOPPM, "-png", "-r", str(dpi), str(pdf), str(out_dir / "slide")],
            check=True, capture_output=True, timeout=60,
        )
        return sorted(out_dir.glob("slide-*.png"))
    except (subprocess.CalledProcessError, subprocess.TimeoutExpired):
        return []


# Session-based storage of uploaded MD files
_SESSIONS: dict[str, Path] = {}


def create_app() -> Flask:
    app = Flask(__name__, template_folder="templates", static_folder="static")
    app.config["MAX_CONTENT_LENGTH"] = 10 * 1024 * 1024  # 10MB

    def _palettes() -> list[str]:
        palettes_dir = Path(__file__).parent.parent / "data" / "themes" / "palettes"
        return sorted(
            p.stem.replace("academic-", "")
            for p in palettes_dir.glob("academic-*.css")
        )

    @app.route("/")
    def index():
        return render_template("index.html", palettes=_palettes())

    @app.route("/editor")
    def editor():
        return render_template("editor.html", palettes=_palettes())

    @app.route("/editor/sample/<name>")
    def editor_sample(name: str):
        """Return a sample Markdown document to populate the editor."""
        templates_dir = Path(__file__).parent.parent / "data" / "templates"
        if name == "minimal":
            md = """---
marp: true
theme: academic
---

<!-- _class: title -->
# 発表タイトル
## サブタイトル
発表者名 / 2026年

---

# 概要
- 背景
- 手法
- 結果
- 考察

---

<!-- _class: end -->
# Thank You
"""
        elif name == "all":
            # Concatenate all 49 templates
            parts = ["---\nmarp: true\ntheme: academic\n---\n"]
            for tpl in sorted(templates_dir.glob("*.md")):
                text = tpl.read_text(encoding="utf-8")
                if text.startswith("---"):
                    end = text.find("---", 3)
                    if end != -1:
                        text = text[end + 3:]
                parts.append(text.strip())
            md = "\n\n---\n\n".join(parts)
        elif name == "academic":
            md = """---
marp: true
theme: academic
---

<!-- _class: title -->
# 研究タイトル
## サブタイトル
山田 太郎 / 2026年4月

---

<!-- _class: agenda -->
# 本日の内容
<div class="agenda-list">
1. 背景と研究目的
2. 提案手法
3. 実験結果
4. 考察とまとめ
</div>

---

<!-- _class: rq -->
# 研究課題
<div class="rq-main">既存手法は大規模データに対してスケールするか？</div>
<div class="rq-sub">計算量 $O(n^2)$ がボトルネックとなっている。</div>

---

<!-- _class: sandwich -->
# 提案手法
<div class="top">
<div class="lead">従来の $O(n^2)$ を $O(n \\log n)$ に改善。</div>
</div>
<div class="columns">
<div>

### 従来
- 計算量: `O(n^2)`
- メモリ: 多い
</div>
<div>

### 提案
- 計算量: `O(n \\log n)`
- メモリ: 少ない
</div>
</div>
<div class="bottom">
<div class="conclusion"><strong>結論:</strong> 大規模データでも実用的な速度を実現。</div>
</div>

---

<!-- _class: kpi -->
# 実験結果
<div class="kpi-container">
<div><span class="kpi-value">97%</span><span class="kpi-label">精度</span></div>
<div><span class="kpi-value">10x</span><span class="kpi-label">高速化</span></div>
<div><span class="kpi-value">50%</span><span class="kpi-label">省メモリ</span></div>
</div>

---

<!-- _class: takeaway -->
# Takeaway
<div class="ta-main">型を選ぶだけで、伝わるプレゼンに</div>
<div class="ta-points">
<ul>
<li>計算量の改善により大規模データに対応</li>
<li>精度は従来と同等</li>
<li>OSS として公開予定</li>
</ul>
</div>

---

<!-- _class: end -->
# Thank You
"""
        else:
            return "unknown sample", 404
        from flask import Response
        return Response(md, mimetype="text/plain; charset=utf-8")

    @app.route("/editor/preview", methods=["POST"])
    def editor_preview():
        """Render MD → PPTX → per-slide PNGs; return list of URLs."""
        from marp_pptx.theme import ThemeConfig, get_default_theme_path, get_palette_path
        from marp_pptx.parser import parse_marp
        from marp_pptx.builder import PptxBuilder

        md_text = request.form.get("markdown", "")
        if not md_text.strip():
            return jsonify({"slides": []})
        palette_name = request.form.get("palette", "")
        try:
            font_scale = float(request.form.get("font_scale", 1.0))
        except ValueError:
            font_scale = 1.0

        # Cache key based on content + settings
        key_src = f"{md_text}|{palette_name}|{font_scale}|math=png".encode("utf-8")
        key = hashlib.md5(key_src).hexdigest()
        out_dir = _PREVIEW_CACHE_DIR / key
        if out_dir.exists():
            pngs = sorted(out_dir.glob("slide-*.png"))
            if pngs:
                return jsonify({"slides": [f"/editor/preview-img/{key}/{p.name}" for p in pngs]})
        out_dir.mkdir(parents=True, exist_ok=True)

        # Build PPTX (expose uploaded images via assets symlink)
        _link_shared_assets_to(out_dir)
        md_path = out_dir / "slides.md"
        md_path.write_text(md_text, encoding="utf-8")
        tc = ThemeConfig.from_css(get_default_theme_path())
        tc.font_scale = max(0.5, min(2.0, font_scale))
        # Force PNG math in the preview: LibreOffice's OMML renderer is
        # unreliable. The download path (_do_convert) keeps OMML for
        # PowerPoint-native editability.
        tc.math_mode = "png"
        if palette_name:
            pp = get_palette_path(palette_name)
            if pp:
                tc.apply_palette(pp)
        slides = parse_marp(str(md_path))
        builder = PptxBuilder(base_path=out_dir, theme=tc)
        builder.build_all(slides)
        pptx_path = out_dir / "slides.pptx"
        builder.save(str(pptx_path))

        # Render to PNGs
        pngs = _render_pptx_to_pngs(pptx_path, out_dir, dpi=90)
        if not pngs:
            return jsonify({"error": "LibreOffice not available or render failed", "slides": []}), 500
        return jsonify({"slides": [f"/editor/preview-img/{key}/{p.name}" for p in pngs]})

    @app.route("/editor/preview-img/<key>/<name>")
    def editor_preview_img(key: str, name: str):
        """Serve a cached preview PNG."""
        if not key.isalnum() or not name.startswith("slide-") or not name.endswith(".png"):
            return "bad path", 400
        png = _PREVIEW_CACHE_DIR / key / name
        if not png.exists():
            return "not found", 404
        return send_file(str(png), mimetype="image/png")

    @app.route("/editor/pptx-to-md", methods=["POST"])
    def editor_pptx_to_md():
        """Upload a PPTX, extract text+structure, return best-effort MD."""
        from marp_pptx.pptx2md import pptx_to_md_with_report

        f = request.files.get("file")
        if not f:
            return jsonify({"error": "no file"}), 400
        tmpdir = Path(tempfile.mkdtemp(prefix="marp_pptx2md_"))
        pptx_path = tmpdir / (f.filename or "input.pptx")
        f.save(str(pptx_path))

        # Extract images into the shared assets dir so the editor can reference them
        assets_dir = _PREVIEW_CACHE_DIR / "shared_assets"
        try:
            report = pptx_to_md_with_report(pptx_path, extract_images_to=assets_dir)
        except Exception as e:
            return jsonify({"error": f"pptx parse failed: {e}"}), 500
        return jsonify(report)

    @app.route("/editor/upload-image", methods=["POST"])
    def editor_upload_image():
        """Receive an image upload, store in shared assets dir, return assets/<name>."""
        f = request.files.get("file")
        if not f:
            return jsonify({"error": "no file"}), 400
        name = (f.filename or "image").lower()
        ext = name.rsplit(".", 1)[-1] if "." in name else ""
        if ext not in ("png", "jpg", "jpeg", "gif", "svg", "webp"):
            return jsonify({"error": "unsupported extension: " + ext}), 400

        upload_dir = _PREVIEW_CACHE_DIR / "shared_assets"
        upload_dir.mkdir(parents=True, exist_ok=True)

        data = f.read()
        digest = hashlib.md5(data).hexdigest()[:10]
        safe_name = f"{digest}_{Path(name).stem[:40]}.{ext}"
        dest = upload_dir / safe_name
        dest.write_bytes(data)
        return jsonify({
            "path": f"assets/{safe_name}",
            "url": f"/editor/asset/{safe_name}",
        })

    @app.route("/editor/asset/<name>")
    def editor_asset_get(name: str):
        if "/" in name or ".." in name:
            return "bad path", 400
        p = _PREVIEW_CACHE_DIR / "shared_assets" / name
        if not p.exists():
            return "not found", 404
        return send_file(str(p))

    def _link_shared_assets_to(out_dir: Path):
        """Symlink uploaded assets into out_dir/assets/ so the builder can resolve
        'assets/foo.png' paths during PPTX generation.
        """
        src = _PREVIEW_CACHE_DIR / "shared_assets"
        if not src.exists():
            return
        dst = out_dir / "assets"
        if dst.exists():
            return
        try:
            dst.symlink_to(src, target_is_directory=True)
        except (OSError, NotImplementedError):
            # fallback: copy (Windows etc.)
            import shutil as _sh
            _sh.copytree(src, dst)

    @app.route("/editor/generate", methods=["POST"])
    def editor_generate():
        """Generate PPTX from raw Markdown text (no file upload)."""
        md_text = request.form.get("markdown", "")
        if not md_text.strip():
            return "empty markdown", 400
        palette_name = request.form.get("palette", "")
        try:
            font_scale = float(request.form.get("font_scale", 1.0))
        except ValueError:
            font_scale = 1.0
        output_name = request.form.get("output_name") or "slides.pptx"

        tmpdir = Path(tempfile.mkdtemp(prefix="marp_editor_"))
        md_path = tmpdir / "slides.md"
        md_path.write_text(md_text, encoding="utf-8")
        return _do_convert(
            md_path=md_path,
            palette_name=palette_name,
            font_scale=font_scale,
            output_name=output_name,
        )

    @app.route("/types-page")
    def types_page():
        from marp_pptx.types import TYPE_REGISTRY, CATEGORIES
        return render_template(
            "types_page.html",
            types=TYPE_REGISTRY,
            categories=CATEGORIES,
        )

    @app.route("/convert", methods=["POST"])
    def convert():
        return _do_convert(
            palette_name=request.form.get("palette", ""),
            font_scale=1.0,
            output_name=None,
        )

    @app.route("/preview", methods=["POST"])
    def preview():
        from marp_pptx.parser import parse_marp
        from marp_pptx.types import get_type_info

        f = request.files.get("file")
        if not f:
            return "No file uploaded", 400

        # Save to session
        session_id = uuid.uuid4().hex
        tmpdir = Path(tempfile.mkdtemp(prefix="marp_preview_"))
        md_path = tmpdir / (f.filename or "slides.md")
        f.save(str(md_path))
        _SESSIONS[session_id] = md_path

        slides_data = parse_marp(str(md_path))
        slides = []
        for sd in slides_data:
            info = get_type_info(sd.slide_class) if sd.slide_class else None
            type_display = sd.slide_class or "default"
            char_count = len(sd.raw)
            bullet_count = sum(
                1 for line in sd.body_lines
                if line.strip().startswith(("- ", "* "))
            )
            table_rows = len(sd.table_rows)
            has_image = bool(sd.image_path) or bool(sd.annotation_figure) or bool(sd.result_figure) or bool(sd.gallery_items)
            has_math = bool(sd.eq_main) or bool(sd.eq_system) or "$" in sd.raw
            warning = None
            if sd.slide_class and not info and sd.slide_class not in (
                "cols-2-wide-l", "cols-2-wide-r",
            ):
                warning = f"未知の型: {sd.slide_class}"
            slides.append({
                "type_display": type_display,
                "h1": sd.h1,
                "h2": sd.h2,
                "char_count": char_count,
                "bullet_count": bullet_count,
                "table_rows": table_rows,
                "has_image": has_image,
                "has_math": has_math,
                "warning": warning,
            })

        filename = md_path.name
        filename_base = md_path.stem

        return render_template(
            "preview.html",
            slides=slides,
            palettes=_palettes(),
            session_id=session_id,
            filename=filename,
            filename_base=filename_base,
        )

    @app.route("/generate", methods=["POST"])
    def generate():
        session_id = request.form.get("session_id", "")
        md_path = _SESSIONS.get(session_id)
        if md_path is None or not md_path.exists():
            return "Session expired. Please re-upload.", 400

        return _do_convert(
            md_path=md_path,
            palette_name=request.form.get("palette", ""),
            font_scale=float(request.form.get("font_scale", 1.0)),
            output_name=request.form.get("output_name") or None,
        )

    def _do_convert(md_path=None, palette_name="", font_scale=1.0, output_name=None):
        from marp_pptx.theme import ThemeConfig, get_default_theme_path, get_palette_path
        from marp_pptx.parser import parse_marp
        from marp_pptx.builder import PptxBuilder

        if md_path is None:
            f = request.files.get("file")
            if not f:
                return "No file uploaded", 400
            tmpdir = Path(tempfile.mkdtemp())
            md_path = tmpdir / (f.filename or "slides.md")
            f.save(str(md_path))

        # Expose uploaded images to the builder's base_path
        _link_shared_assets_to(md_path.parent)

        tc = ThemeConfig.from_css(get_default_theme_path())
        tc.font_scale = max(0.5, min(2.0, font_scale))
        if palette_name:
            pp = get_palette_path(palette_name)
            if pp:
                tc.apply_palette(pp)

        slides = parse_marp(str(md_path))
        builder = PptxBuilder(base_path=md_path.parent, theme=tc)
        builder.build_all(slides)

        out_name = output_name or (md_path.stem + "_editable.pptx")
        out_path = md_path.parent / out_name
        builder.save(str(out_path))

        return send_file(
            str(out_path),
            as_attachment=True,
            download_name=out_name,
            mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )

    @app.route("/api/types")
    def api_types():
        from marp_pptx.types import TYPE_REGISTRY, CATEGORIES
        data = [
            {
                "name": t.name,
                "css_class": t.css_class,
                "category": t.category,
                "category_ja": CATEGORIES.get(t.category, t.category),
                "geometry": t.geometry,
                "meaning": t.meaning,
                "use_when": t.use_when,
            }
            for t in TYPE_REGISTRY
        ]
        return jsonify(data)

    return app
