"""Smoke tests for the Flask web UI.

These guard the template/static externalization: every GET route must still
render, the editor must expose its key element IDs, and the JSON/type APIs must
stay intact. They intentionally assert on structural markers (titles, element
IDs, type names) rather than CSS/JS so they survive asset extraction.
"""
import io
import json

import pytest

pytest.importorskip("flask")

from marp_pptx.web.app import create_app


@pytest.fixture
def client():
    app = create_app()
    app.config["TESTING"] = True
    return app.test_client()


def test_index(client):
    r = client.get("/")
    assert r.status_code == 200
    assert b"marp-pptx Web UI" in r.data


def test_editor_has_key_elements(client):
    r = client.get("/editor")
    assert r.status_code == 200
    body = r.data
    assert b"marp-pptx Editor" in body
    for marker in (b'id="md-editor"', b'id="preview-content"', b'id="md-upload"'):
        assert marker in body, marker


def test_types_page_lists_types(client):
    r = client.get("/types-page")
    assert r.status_code == 200
    assert b"marp-pptx: Slide Types" in r.data
    # a couple of known type names must appear
    assert b"funnel" in r.data
    assert b"kpi" in r.data


def test_api_types(client):
    r = client.get("/api/types")
    assert r.status_code == 200
    data = json.loads(r.data)
    assert isinstance(data, list)
    assert len(data) == 52
    names = {t["name"] for t in data}
    assert {"title", "funnel", "statement", "big-number", "chart"} <= names


def test_sample_markdown(client):
    r = client.get("/editor/sample/minimal")
    assert r.status_code == 200
    assert b"marp: true" in r.data
    assert b"_class: title" in r.data


def test_external_css_served(client):
    # editor links its stylesheet rather than inlining it, and the file serves
    body = client.get("/editor").data
    assert b"css/editor.css" in body
    assert b"<style>" not in body
    css = client.get("/static/css/editor.css")
    assert css.status_code == 200
    assert b"{" in css.data  # looks like CSS


def test_external_js_served(client):
    # editor links its script rather than inlining it, and the file serves
    body = client.get("/editor").data
    assert b"js/editor.js" in body
    assert b"<script>" not in body  # no inline script block remains
    js = client.get("/static/js/editor.js")
    assert js.status_code == 200
    assert b"function" in js.data


def test_preview_breakdown(client):
    md = (
        "---\nmarp: true\n---\n\n"
        "<!-- _class: title -->\n# T\n## S\nme\n\n---\n\n"
        "<!-- _class: kpi -->\n# K\n"
        '<div class="kpi-container"><div>'
        '<span class="kpi-value">9</span><span class="kpi-label">x</span>'
        "</div></div>\n"
    )
    r = client.post(
        "/preview",
        data={"file": (io.BytesIO(md.encode()), "deck.md")},
        content_type="multipart/form-data",
    )
    assert r.status_code == 200
    # the breakdown table names each slide's type
    assert b"title" in r.data
    assert b"kpi" in r.data
