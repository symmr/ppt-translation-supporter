const test = require("node:test");
const assert = require("node:assert/strict");
const JSZip = require("jszip");
const {
  extractDocxTexts,
  injectDocxTexts,
  listDocxRasterPictures,
  replaceDocxPictureTextBoxes,
} = require("../docs/docx-text.js");

const NS = {
  w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
  r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  a: "http://schemas.openxmlformats.org/drawingml/2006/main",
  wp: "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
  pic: "http://schemas.openxmlformats.org/drawingml/2006/picture",
  wps: "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
};

function documentXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${NS.w}" xmlns:r="${NS.r}" xmlns:a="${NS.a}" xmlns:wp="${NS.wp}" xmlns:pic="${NS.pic}" xmlns:wps="${NS.wps}">
  <w:body>
    <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Title</w:t></w:r></w:p>
    <w:p><w:r><w:t>Hello</w:t></w:r><w:r><w:t> world</w:t></w:r></w:p>
    <w:tbl><w:tr><w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
    <w:p><w:r><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
      <mc:Choice Requires="wps">
        <w:drawing><wp:anchor>
          <wp:extent cx="100" cy="40"/>
          <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
            <wps:wsp><wps:cNvSpPr txBox="1"/>
              <wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="40"/></a:xfrm></wps:spPr>
              <wps:txbx><w:txbxContent><w:p><w:r><w:t>Box</w:t></w:r></w:p></w:txbxContent></wps:txbx>
              <wps:bodyPr wrap="square"/>
            </wps:wsp>
          </a:graphicData></a:graphic>
        </wp:anchor></w:drawing>
      </mc:Choice>
      <mc:Fallback><w:p><w:r><w:t>Box</w:t></w:r></w:p></mc:Fallback>
    </mc:AlternateContent></w:r></w:p>
    <w:p><w:r><w:drawing><wp:inline>
      <wp:extent cx="200" cy="100"/>
      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
        <pic:pic>
          <pic:nvPicPr><pic:cNvPr id="3" name="Picture"/><pic:cNvPicPr/></pic:nvPicPr>
          <pic:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
          <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="200" cy="100"/></a:xfrm></pic:spPr>
        </pic:pic>
      </a:graphicData></a:graphic>
    </wp:inline></w:drawing></w:r></w:p>
    <w:p><w:r><w:drawing><wp:inline>
      <wp:extent cx="50" cy="50"/>
      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
        <pic:pic>
          <pic:blipFill><a:blip r:embed="rId2"/></pic:blipFill>
          <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="50" cy="50"/></a:xfrm></pic:spPr>
        </pic:pic>
      </a:graphicData></a:graphic>
    </wp:inline></w:drawing></w:r></w:p>
  </w:body>
</w:document>`;
}

async function buildZip() {
  const zip = new JSZip();
  zip.file("word/document.xml", documentXml());
  zip.file("word/_rels/document.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image2.emf"/>
</Relationships>`);
  zip.file("word/media/image1.png", Buffer.from([137, 80, 78, 71]));
  return zip;
}

test("extracts body, table and text box once, skipping the VML fallback", async () => {
  const extracted = await extractDocxTexts(await buildZip());
  assert.deepEqual(extracted.lines, [
    "uid_0001",
    "Title",
    "uid_0002",
    "[0]Hello[/0][1] world[/1]",
    "uid_0003",
    "Cell",
    "uid_0004",
    "Box",
  ]);
  assert.equal(extracted.metadata[0].title, true);
  assert.equal(extracted.metadata[1].title, false);
  assert.equal(extracted.metadata[2].type, "table");
  assert.equal(extracted.metadata[3].type, "textbox");
});

test("inject writes translated runs and leaves the picture in place", async () => {
  const zip = await buildZip();
  const extracted = await extractDocxTexts(zip);
  const translations = {
    uid_0001: "題名",
    uid_0002: "[0]こんにちは[/0][1] 世界[/1]",
    uid_0003: "セル",
    uid_0004: "箱",
  };
  const result = await injectDocxTexts(zip, translations, extracted.metadata, { bodyFont: "Noto Sans JP" });
  assert.equal(result.injected, 4);
  const again = await extractDocxTexts(zip);
  assert.match(again.text, /題名/);
  assert.match(again.text, /こんにちは/);
  assert.match(again.text, /世界/);
  const xml = await zip.file("word/document.xml").async("string");
  assert.match(xml, /r:embed="rId1"/);
  assert.match(xml, /eastAsia="Noto Sans JP"/);
});

test("lists raster pictures and replaces text boxes on that picture only", async () => {
  const zip = await buildZip();
  const pictures = await listDocxRasterPictures(zip);
  assert.deepEqual(pictures.map((pic) => pic.media), ["word/media/image1.png"]);
  assert.equal(pictures[0].cx, 200);
  assert.equal(pictures[0].cy, 100);

  const first = await replaceDocxPictureTextBoxes(zip, pictures[0].slidePath, "imgtext-1-1-", [
    { x: 10, y: 20, cx: 80, cy: 30, text: "図の文字", fill: "112233", ink: "FFFFFF" },
  ]);
  assert.equal(first.added, 1);
  let xml = await zip.file("word/document.xml").async("string");
  assert.match(xml, /wordprocessingGroup/);
  assert.match(xml, /図の文字/);
  assert.match(xml, /wrap="square"/);
  assert.match(xml, /<a:noAutofit\/>/);
  assert.match(xml, /<w:txbxContent><w:p><w:pPr>/);
  assert.match(xml, /w:lineRule="exact"/);
  assert.doesNotMatch(xml, /<w:vanish\/>/);

  const extracted = await extractDocxTexts(zip);
  assert.match(extracted.text, /図の文字/);
  assert.doesNotMatch(extracted.text, /imgtext-1-1-1/);

  const second = await replaceDocxPictureTextBoxes(zip, pictures[0].slidePath, "imgtext-1-1-", [
    { x: 10, y: 20, cx: 80, cy: 30, text: "差し替え", fill: "FFFFFF", ink: "222222" },
  ]);
  assert.equal(second.removed, 1);
  assert.equal(second.added, 1);
  xml = await zip.file("word/document.xml").async("string");
  assert.match(xml, /差し替え/);
  assert.doesNotMatch(xml, /図の文字/);
  const after = await extractDocxTexts(zip);
  assert.match(after.text, /Box/);
});
