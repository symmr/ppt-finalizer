"use strict";

const { test, describe } = require("node:test");
const assert = require("node:assert/strict");
const JSZip = require("jszip");

const core = require("../docs/pptx-core.js");

const REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

function rel(id, type, target) {
  return `<Relationship Id="${id}" Type="${REL_NS}/${type}" Target="${target}"/>`;
}

function relsXml(rels) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `${rels.join("")}</Relationships>`;
}

// Builds a minimal but structurally realistic PPTX package in memory:
// - slide1 uses slideLayout1 -> slideMaster1, and embeds media/image1.png
// - slideLayout2 is unused by any slide but still embeds media/image3.png
// - slideMaster2 is unused by any layout
// - media/image2.png is referenced nowhere (true package orphan)
// - slide1 has a notes slide, and there is one notes master
async function buildFixtureZip() {
  const zip = new JSZip();

  zip.file(
    "ppt/presentation.xml",
    `<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
    `<p:sldMasterIdLst><p:sldMasterId id="1" r:id="rId3"/><p:sldMasterId id="2" r:id="rId5"/></p:sldMasterIdLst>` +
    `<p:notesMasterIdLst><p:notesMasterId r:id="rId6"/></p:notesMasterIdLst>` +
    `<p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst>` +
    `</p:presentation>`
  );
  zip.file(
    "ppt/_rels/presentation.xml.rels",
    relsXml([
      rel("rId2", "slide", "slides/slide1.xml"),
      rel("rId3", "slideMaster", "slideMasters/slideMaster1.xml"),
      rel("rId5", "slideMaster", "slideMasters/slideMaster2.xml"),
      rel("rId6", "notesMaster", "notesMasters/notesMaster1.xml"),
    ])
  );

  zip.file("ppt/slides/slide1.xml", `<p:sld><p:cSld><p:spTree/></p:cSld></p:sld>`);
  zip.file(
    "ppt/slides/_rels/slide1.xml.rels",
    relsXml([
      rel("rId1", "slideLayout", "../slideLayouts/slideLayout1.xml"),
      rel("rId2", "image", "../media/image1.png"),
      rel("rId3", "notesSlide", "../notesSlides/notesSlide1.xml"),
    ])
  );

  zip.file("ppt/slideLayouts/slideLayout1.xml", `<p:sldLayout/>`);
  zip.file(
    "ppt/slideLayouts/_rels/slideLayout1.xml.rels",
    relsXml([rel("rId1", "slideMaster", "../slideMasters/slideMaster1.xml")])
  );

  zip.file("ppt/slideLayouts/slideLayout2.xml", `<p:sldLayout/>`);
  zip.file(
    "ppt/slideLayouts/_rels/slideLayout2.xml.rels",
    relsXml([
      rel("rId1", "slideMaster", "../slideMasters/slideMaster1.xml"),
      rel("rId2", "image", "../media/image3.png"),
    ])
  );

  zip.file("ppt/slideMasters/slideMaster1.xml", `<p:sldMaster/>`);
  zip.file("ppt/slideMasters/slideMaster2.xml", `<p:sldMaster/>`);

  zip.file("ppt/media/image1.png", "used-directly");
  zip.file("ppt/media/image2.png", "package-orphan");
  zip.file("ppt/media/image3.png", "used-only-by-unused-layout");

  zip.file("ppt/notesSlides/notesSlide1.xml", `<p:notes/>`);
  zip.file("ppt/notesMasters/notesMaster1.xml", `<p:notesMaster/>`);

  zip.file(
    "docProps/core.xml",
    `<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/">` +
    `<dc:title>Test Deck</dc:title><dc:creator>Alice</dc:creator><cp:revision>5</cp:revision>` +
    `</cp:coreProperties>`
  );
  zip.file(
    "docProps/app.xml",
    `<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">` +
    `<Company>Acme</Company></Properties>`
  );

  // Round-trip through generate/load so entries carry real JSZip metadata
  // (compressed/uncompressed sizes), matching how the app loads a real file.
  const buffer = await zip.generateAsync({ type: "nodebuffer" });
  return JSZip.loadAsync(buffer);
}

describe("getPresentationSlideOrder", () => {
  test("maps the single slide to display order 1", async () => {
    const zip = await buildFixtureZip();
    const order = await core.getPresentationSlideOrder(zip);
    assert.equal(order.get("ppt/slides/slide1.xml"), 1);
  });
});

describe("buildDeckStructure / computeCleanupPlan", () => {
  test("flags the unused layout and unused master", async () => {
    const zip = await buildFixtureZip();
    const plan = await core.computeCleanupPlan(zip);

    assert.deepEqual(plan.layoutsToRemove, ["ppt/slideLayouts/slideLayout2.xml"]);
    assert.deepEqual(plan.mastersToRemove, ["ppt/slideMasters/slideMaster2.xml"]);
    assert.equal(plan.unusedLayoutCount, 1);
    assert.equal(plan.unusedMasterCount, 1);
  });

  test("media only reachable through the unused layout is reported as freed by structure cleanup", async () => {
    const zip = await buildFixtureZip();
    const plan = await core.computeCleanupPlan(zip);
    const freedPaths = plan.structureFreedMedia.items.map((item) => item.path);
    assert.deepEqual(freedPaths, ["ppt/media/image3.png"]);
  });

  test("notes and document properties are detected", async () => {
    const zip = await buildFixtureZip();
    const plan = await core.computeCleanupPlan(zip);
    assert.equal(plan.notes.count, 1);
    assert.equal(plan.properties.fieldCount, 4); // title, creator, revision, company
  });
});

describe("computePackageOrphanMedia", () => {
  test("only the never-referenced media file counts as a package orphan", async () => {
    const zip = await buildFixtureZip();
    const orphans = await core.computePackageOrphanMedia(zip);
    const paths = orphans.items.map((item) => item.path);
    assert.deepEqual(paths, ["ppt/media/image2.png"]);
  });
});

describe("analyzeMediaUsage", () => {
  test("labels directly-used media with its slide number", async () => {
    const zip = await buildFixtureZip();
    const slidePathToNum = await core.getPresentationSlideOrder(zip);
    const analysis = await core.analyzeMediaUsage(zip, slidePathToNum);
    const image1 = analysis.items.find((item) => item.path === "ppt/media/image1.png");
    assert.deepEqual(image1.slides, ["スライド 1（直接）"]);
    assert.equal(image1.isOrphan, false);
  });

  test("labels media reachable only through an unused layout as orphaned", async () => {
    const zip = await buildFixtureZip();
    const slidePathToNum = await core.getPresentationSlideOrder(zip);
    const analysis = await core.analyzeMediaUsage(zip, slidePathToNum);
    const image3 = analysis.items.find((item) => item.path === "ppt/media/image3.png");
    assert.equal(image3.isOrphan, true);
    assert.match(image3.slides[0], /^未使用レイアウト/);
  });

  test("media with no relationship at all has no usage labels", async () => {
    const zip = await buildFixtureZip();
    const slidePathToNum = await core.getPresentationSlideOrder(zip);
    const analysis = await core.analyzeMediaUsage(zip, slidePathToNum);
    const image2 = analysis.items.find((item) => item.path === "ppt/media/image2.png");
    assert.deepEqual(image2.slides, []);
    assert.equal(image2.isOrphan, false); // "no usage info" differs from "confirmed orphan"
  });
});

describe("removeNotesFromZip", () => {
  test("deletes the notes slide, notes master, and their relationships", async () => {
    const zip = await buildFixtureZip();
    const result = await core.removeNotesFromZip(zip);

    assert.equal(result.count, 1);
    assert.equal(zip.files["ppt/notesSlides/notesSlide1.xml"], undefined);
    assert.equal(zip.files["ppt/notesMasters/notesMaster1.xml"], undefined);

    const slideRels = await zip.files["ppt/slides/_rels/slide1.xml.rels"].async("string");
    assert.doesNotMatch(slideRels, /notesSlide/);

    const presXml = await zip.files["ppt/presentation.xml"].async("string");
    assert.doesNotMatch(presXml, /notesMasterIdLst/);
  });
});

describe("scanDocumentProperties / clearDocumentProperties", () => {
  test("scan reports the populated fields", async () => {
    const zip = await buildFixtureZip();
    const props = await core.scanDocumentProperties(zip);
    const labels = props.fields.map((f) => f.label);
    assert.ok(labels.includes("タイトル"));
    assert.ok(labels.includes("作成者"));
    assert.ok(labels.includes("会社"));
  });

  test("clear empties the known fields and resets revision to 0", async () => {
    const zip = await buildFixtureZip();
    await core.clearDocumentProperties(zip);

    const coreXml = await zip.files["docProps/core.xml"].async("string");
    assert.equal(core.readXmlElementText(coreXml, "dc:title"), "");
    assert.equal(core.readXmlElementText(coreXml, "cp:revision"), "0");

    const appXml = await zip.files["docProps/app.xml"].async("string");
    assert.equal(core.readXmlElementText(appXml, "Company"), "");
  });
});

describe("removeUnusedStructure", () => {
  test("removes the unused layout/master parts and detaches them from their owners", async () => {
    const zip = await buildFixtureZip();
    const plan = await core.computeCleanupPlan(zip);
    const result = await core.removeUnusedStructure(zip, plan);

    assert.equal(result.layoutsRemoved, 1);
    assert.equal(result.mastersRemoved, 1);
    assert.equal(zip.files["ppt/slideLayouts/slideLayout2.xml"], undefined);
    assert.equal(zip.files["ppt/slideMasters/slideMaster2.xml"], undefined);

    const presRels = await zip.files["ppt/_rels/presentation.xml.rels"].async("string");
    assert.doesNotMatch(presRels, /slideMaster2\.xml/);
  });
});
