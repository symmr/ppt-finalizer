"use strict";

const { test, describe } = require("node:test");
const assert = require("node:assert/strict");

const core = require("../docs/pptx-core.js");

describe("XML entity helpers", () => {
  test("decodeXmlEntities decodes the five predefined entities", () => {
    assert.equal(
      core.decodeXmlEntities("&amp;&lt;&gt;&quot;&apos;"),
      "&<>\"'"
    );
  });

  test("escapeXmlAttr / decodeXmlEntities round-trip arbitrary text", () => {
    const original = `Fira Code & "Bold" <Italic> 'Test'`;
    const escaped = core.escapeXmlAttr(original);
    assert.equal(core.decodeXmlEntities(escaped), original);
  });

  test("escapeHtml escapes all five characters", () => {
    assert.equal(
      core.escapeHtml(`<b>a & b "c" 'd'</b>`),
      "&lt;b&gt;a &amp; b &quot;c&quot; 'd'&lt;/b&gt;"
    );
  });
});

describe("formatBytes / formatSizeChange", () => {
  test("formats bytes, kilobytes, and megabytes", () => {
    assert.equal(core.formatBytes(500), "500 B");
    assert.equal(core.formatBytes(2048), "2.0 KB");
    assert.equal(core.formatBytes(5 * 1024 * 1024), "5.00 MB");
  });

  test("reports a reduction when the file shrinks", () => {
    assert.equal(core.formatSizeChange(1000, 800), "200 B 削減（−20.0%）");
  });

  test("reports an increase when the file grows", () => {
    assert.equal(core.formatSizeChange(800, 1000), "200 B 増加（+25.0%）");
  });

  test("reports no change when sizes are equal", () => {
    assert.equal(core.formatSizeChange(500, 500), "変化なし（0%）");
  });
});

describe("parseRelationships", () => {
  test("extracts id/type/target from self-closing and closed elements", () => {
    const xml = `<?xml version="1.0"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
        <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png" ></Relationship>
      </Relationships>`;
    const rels = core.parseRelationships(xml);
    assert.equal(rels.length, 2);
    assert.equal(rels[0].id, "rId1");
    assert.match(rels[0].type, /slideLayout$/);
    assert.equal(rels[0].target, "../slideLayouts/slideLayout1.xml");
    assert.equal(rels[1].target, "../media/image1.png");
  });
});

describe("resolveZipPath", () => {
  test("resolves relative paths against the owner's directory", () => {
    assert.equal(
      core.resolveZipPath("ppt/slides/slide1.xml", "../media/image1.png"),
      "ppt/media/image1.png"
    );
  });

  test("treats a leading slash as package-root relative", () => {
    assert.equal(
      core.resolveZipPath("ppt/slides/slide1.xml", "/ppt/media/image1.png"),
      "ppt/media/image1.png"
    );
  });

  test("collapses redundant '..' segments", () => {
    assert.equal(
      core.resolveZipPath("ppt/slideLayouts/slideLayout1.xml", "../slideMasters/../slideMasters/slideMaster1.xml"),
      "ppt/slideMasters/slideMaster1.xml"
    );
  });
});

describe("isMediaRelationship", () => {
  test("is true when the target path lives under a media/ folder", () => {
    assert.equal(core.isMediaRelationship("some/custom/type", "../media/image1.png"), true);
  });

  test("is true when the relationship type carries a known media marker", () => {
    assert.equal(
      core.isMediaRelationship(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
        "../embeddings/oleObject1.bin"
      ),
      true
    );
  });

  test("is false for unrelated relationship types/targets", () => {
    assert.equal(
      core.isMediaRelationship(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout",
        "../slideLayouts/slideLayout1.xml"
      ),
      false
    );
  });
});

describe("replaceFontsContextual", () => {
  function shape(phType, typeface) {
    const ph = phType ? `<p:ph type="${phType}"/>` : "<p:ph/>";
    return `<p:sp><p:nvSpPr><p:nvPr>${ph}</p:nvPr></p:nvSpPr>` +
      `<p:txBody><a:r><a:rPr><a:latin typeface="${typeface}"/></a:rPr></a:r></p:txBody></p:sp>`;
  }

  test("gives the title font to title/ctrTitle/subTitle placeholders", () => {
    const xml = shape("title", "Original");
    const { data, count } = core.replaceFontsContextual(xml, "TitleFont", "BodyFont");
    assert.match(data, /typeface="TitleFont"/);
    assert.equal(count, 1);
  });

  test("gives the body font to body placeholders", () => {
    const xml = shape("body", "Original");
    const { data } = core.replaceFontsContextual(xml, "TitleFont", "BodyFont");
    assert.match(data, /typeface="BodyFont"/);
  });

  test("falls back to the body font for shapes without a title/body placeholder", () => {
    const xml = shape(null, "Original");
    const { data } = core.replaceFontsContextual(xml, "TitleFont", "BodyFont");
    assert.match(data, /typeface="BodyFont"/);
  });

  test("routes <p:titleStyle>/<p:bodyStyle> blocks in slide masters by role", () => {
    const xml =
      `<p:txStyles>` +
      `<p:titleStyle><a:lvl1pPr><a:defRPr><a:latin typeface="Original"/></a:defRPr></a:lvl1pPr></p:titleStyle>` +
      `<p:bodyStyle><a:lvl1pPr><a:defRPr><a:latin typeface="Original"/></a:defRPr></a:lvl1pPr></p:bodyStyle>` +
      `</p:txStyles>`;
    const { data } = core.replaceFontsContextual(xml, "TitleFont", "BodyFont");
    const titleBlock = data.match(/<p:titleStyle>[\s\S]*?<\/p:titleStyle>/)[0];
    const bodyBlock = data.match(/<p:bodyStyle>[\s\S]*?<\/p:bodyStyle>/)[0];
    assert.match(titleBlock, /typeface="TitleFont"/);
    assert.match(bodyBlock, /typeface="BodyFont"/);
  });

  // Known limitation (see project review): the replacement is not scoped to
  // <a:latin>/<a:ea>/<a:cs>, so a bullet's symbol font (e.g. Wingdings) is
  // overwritten too, which can break bullet glyphs. Documented here rather
  // than fixed, per explicit decision to leave this for a follow-up change.
  test("also rewrites <a:buFont> bullet typefaces (known bug, not yet fixed)", () => {
    const xml = `<p:sp><p:nvSpPr><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr>` +
      `<a:pPr><a:buFont typeface="Wingdings"/></a:pPr></p:sp>`;
    const { data } = core.replaceFontsContextual(xml, "TitleFont", "BodyFont");
    assert.match(data, /<a:buFont typeface="BodyFont"\/>/);
  });
});
