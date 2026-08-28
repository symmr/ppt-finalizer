"use strict";

// Regenerates test/fixtures/smoke-test.pptx: a small but structurally real
// PPTX used for manual/browser smoke testing (drag it into the app) and as
// input for ad-hoc scripts. Not part of `npm test` — the automated tests use
// in-memory JSZip fixtures instead. Run with: node test/fixtures/build-smoke-pptx.js
//
// Contents, by design:
// - 1 slide with a title placeholder (typeface Arial) and a body placeholder
//   (typeface Calibri) -> exercises title/body font replacement
// - slideLayout1 (used) and slideLayout2 (unused) sharing slideMaster1
//   -> exercises unused-layout detection
// - ppt/media/image1.png referenced directly by the slide (kept)
// - ppt/media/image2.png referenced nowhere (orphan, removable)
// - docProps/core.xml + app.xml with a title/creator/company set
//   -> exercises document property scanning/clearing

const path = require("path");
const fs = require("fs");
const JSZip = require("jszip");

const REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const rel = (id, type, target) => `<Relationship Id="${id}" Type="${REL_NS}/${type}" Target="${target}"/>`;
const relsXml = (rels) =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join("")}</Relationships>`;

// 1x1 transparent PNG
const PNG_1X1 = Buffer.from(
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=",
  "base64"
);

async function build() {
  const zip = new JSZip();

  zip.file(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
    `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
    `<Default Extension="xml" ContentType="application/xml"/>` +
    `<Default Extension="png" ContentType="image/png"/>` +
    `<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>` +
    `</Types>`
  );
  zip.file("_rels/.rels", relsXml([rel("rId1", "officeDocument", "ppt/presentation.xml")]));

  zip.file(
    "ppt/presentation.xml",
    `<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
    `<p:sldMasterIdLst><p:sldMasterId id="1" r:id="rId3"/></p:sldMasterIdLst>` +
    `<p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst>` +
    `</p:presentation>`
  );
  zip.file(
    "ppt/_rels/presentation.xml.rels",
    relsXml([
      rel("rId2", "slide", "slides/slide1.xml"),
      rel("rId3", "slideMaster", "slideMasters/slideMaster1.xml"),
    ])
  );

  zip.file(
    "ppt/slides/slide1.xml",
    `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree>` +
    `<p:sp><p:nvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:txBody><a:p><a:r><a:rPr><a:latin typeface="Arial"/></a:rPr><a:t>Hello</a:t></a:r></a:p></p:txBody></p:sp>` +
    `<p:sp><p:nvSpPr><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr><p:txBody><a:p><a:r><a:rPr><a:latin typeface="Calibri"/></a:rPr><a:t>World</a:t></a:r></a:p></p:txBody></p:sp>` +
    `</p:spTree></p:cSld></p:sld>`
  );
  zip.file(
    "ppt/slides/_rels/slide1.xml.rels",
    relsXml([
      rel("rId1", "slideLayout", "../slideLayouts/slideLayout1.xml"),
      rel("rId2", "image", "../media/image1.png"),
    ])
  );

  zip.file("ppt/slideLayouts/slideLayout1.xml", "<p:sldLayout/>");
  zip.file(
    "ppt/slideLayouts/_rels/slideLayout1.xml.rels",
    relsXml([rel("rId1", "slideMaster", "../slideMasters/slideMaster1.xml")])
  );
  zip.file("ppt/slideLayouts/slideLayout2.xml", "<p:sldLayout/>");
  zip.file(
    "ppt/slideLayouts/_rels/slideLayout2.xml.rels",
    relsXml([rel("rId1", "slideMaster", "../slideMasters/slideMaster1.xml")])
  );

  zip.file("ppt/slideMasters/slideMaster1.xml", "<p:sldMaster/>");

  zip.file("ppt/media/image1.png", PNG_1X1);
  zip.file("ppt/media/image2.png", PNG_1X1);

  zip.file(
    "docProps/core.xml",
    `<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/">` +
    `<dc:title>Smoke Test Deck</dc:title><dc:creator>QA Bot</dc:creator></cp:coreProperties>`
  );
  zip.file(
    "docProps/app.xml",
    `<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Company>Acme</Company></Properties>`
  );

  const buffer = await zip.generateAsync({ type: "nodebuffer" });
  const outPath = path.join(__dirname, "smoke-test.pptx");
  fs.writeFileSync(outPath, buffer);
  console.log(`wrote ${buffer.length} bytes to ${outPath}`);
}

build();
