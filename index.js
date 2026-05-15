const express = require('express');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, BorderStyle, WidthType, ShadingType, LevelFormat,
  AlignmentType, PageNumber, Header, Footer,
  TabStopType, TabStopPosition
} = require('docx');
const fs = require('fs');

const app = express();
app.use(express.json({ limit: '10mb' }));

// ── Palette ────────────────────────────────────────────────────────
const C = {
  navy:    "1A2B4A",
  gold:    "C8922A",
  blue:    "2E5090",
  body:    "222222",
  lgray:   "888888",
  mgray:   "555555",
  altrow:  "F2F5FB",
  white:   "FFFFFF",
  p0bg:    "FDF0EF",  p0txt: "B02A20",
  p1bg:    "FEF6EC",  p1txt: "B05C10",
  p2bg:    "EEF8EE",  p2txt: "1A6B1A",
  veribg:  "EAF5EE",  veritxt: "145C2E",
  calcbg:  "E8EFF8",  calctxt: "1A4A88",
  assmbg:  "FDF6E3",  asmtxt:  "7A5900",
  needbg:  "FDF1EE",  neetxt:  "8A2A1A",
};

// Content width: 12240 - 1080 - 1080 = 10080 DXA (0.75" margins each side)
const PAGE_W    = 12240;
const MARGIN    = 1080;
const CONTENT_W = PAGE_W - MARGIN * 2; // 10080

// ── URL truncator ─────────────────────────────────────────────────
function truncateUrl(text) {
  const s = String(text ?? "");
  // If it looks like a full URL, shorten to path only, max 40 chars
  if (s.startsWith("http://") || s.startsWith("https://")) {
    try {
      const u = new URL(s);
      const path = u.pathname.length > 40
        ? "…" + u.pathname.slice(-38)
        : u.pathname || "/";
      return u.hostname.replace("www.", "") + path;
    } catch (_) { /* fall through */ }
  }
  return s;
}

// ── Priority / Status cell styling ────────────────────────────────
function cellStyle(text) {
  const v = String(text ?? "").trim();
  if (v === "P0") return { bg: C.p0bg, txt: C.p0txt, bold: true };
  if (v === "P1") return { bg: C.p1bg, txt: C.p1txt, bold: true };
  if (v === "P2") return { bg: C.p2bg, txt: C.p2txt, bold: true };
  if (v.includes("Verified Data"))        return { bg: C.veribg, txt: C.veritxt, bold: false };
  if (v.includes("Calculated Estimate"))  return { bg: C.calcbg, txt: C.calctxt, bold: false };
  if (v.includes("Strategic Assumption")) return { bg: C.assmbg, txt: C.asmtxt,  bold: false };
  if (v.includes("Needs Verification"))   return { bg: C.needbg, txt: C.neetxt,  bold: false };
  return null;
}

// ── Smart column width distribution ───────────────────────────────
// Assigns widths based on column count and typical content.
// First col gets more space (labels), remaining split proportionally.
function getColWidths(colCount) {
  if (colCount === 1) return [CONTENT_W];

  // Weight hints: col 0 is always a label/name col, gets more room.
  // For very wide tables (7+ cols) compress everything evenly.
  const total = CONTENT_W;

  if (colCount === 2) {
    return [Math.round(total * 0.38), Math.round(total * 0.62)];
  }
  if (colCount === 3) {
    const a = Math.round(total * 0.28);
    const b = Math.round(total * 0.30);
    const c = total - a - b;
    return [a, b, c];
  }
  if (colCount === 4) {
    const a = Math.round(total * 0.25);
    const rest = total - a;
    const each = Math.round(rest / 3);
    return [a, each, each, total - a - each * 2];
  }
  if (colCount === 5) {
    const a = Math.round(total * 0.22);
    const rest = total - a;
    const each = Math.round(rest / 4);
    return [a, each, each, each, total - a - each * 3];
  }
  if (colCount === 6) {
    const a = Math.round(total * 0.20);
    const rest = total - a;
    const each = Math.round(rest / 5);
    return [a, each, each, each, each, total - a - each * 4];
  }
  // 7+ cols: equal distribution, slight boost for first col
  const a = Math.round(total * 0.17);
  const rest = total - a;
  const each = Math.round(rest / (colCount - 1));
  const widths = [a];
  for (let i = 1; i < colCount - 1; i++) widths.push(each);
  widths.push(total - a - each * (colCount - 2));
  return widths;
}

// ── Table cell builder ─────────────────────────────────────────────
function makeCell(rawText, isHeader, width, rowIdx) {
  const border = { style: BorderStyle.SINGLE, size: 1, color: "D0D8E8" };
  const topBorder = isHeader
    ? { style: BorderStyle.SINGLE, size: 4, color: C.gold }
    : border;
  const botBorder = isHeader
    ? { style: BorderStyle.SINGLE, size: 4, color: C.gold }
    : border;

  const text = truncateUrl(rawText);
  const style = !isHeader ? cellStyle(text) : null;

  const txtColor = isHeader ? C.white : (style ? style.txt : C.body);
  const isBold   = isHeader || (style ? style.bold : false);
  const bgFill   = isHeader
    ? C.navy
    : (style ? style.bg : (rowIdx % 2 === 1 ? C.altrow : C.white));

  return new TableCell({
    borders: { top: topBorder, bottom: botBorder, left: border, right: border },
    width: { size: width, type: WidthType.DXA },
    shading: { fill: bgFill, type: ShadingType.CLEAR },
    margins: { top: 90, bottom: 90, left: 120, right: 120 },
    verticalAlign: "center",
    children: [new Paragraph({
      alignment: isHeader ? AlignmentType.CENTER : AlignmentType.LEFT,
      spacing: { before: 0, after: 0 },
      children: [new TextRun({
        text,
        font: "Arial",
        size: 17,       // 8.5pt — readable without ballooning row height
        bold: isBold,
        color: txtColor
      })]
    })]
  });
}

// ── Table builder ─────────────────────────────────────────────────
function buildTable(headers, rows) {
  const colWidths = getColWidths(headers.length);

  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: colWidths,
    rows: [
      new TableRow({
        tableHeader: true,
        children: headers.map((h, i) => makeCell(h, true, colWidths[i], 0))
      }),
      ...rows.map((r, rowIdx) => new TableRow({
        children: r.map((c, i) => makeCell(c, false, colWidths[i], rowIdx))
      }))
    ]
  });
}

// ── Callout table (single-row highlight box) ──────────────────────
// Used for type:"callout" — visually distinct bordered box
function buildCallout(headers, row) {
  const colWidths = getColWidths(headers.length);
  const goldBorder = { style: BorderStyle.SINGLE, size: 6, color: C.gold };
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: colWidths,
    rows: [
      new TableRow({
        children: headers.map((h, i) => new TableCell({
          borders: { top: goldBorder, bottom: { style: BorderStyle.SINGLE, size: 1, color: "D0D8E8" }, left: goldBorder, right: goldBorder },
          width: { size: colWidths[i], type: WidthType.DXA },
          shading: { fill: "F0F4FA", type: ShadingType.CLEAR },
          margins: { top: 80, bottom: 60, left: 120, right: 120 },
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 0, after: 0 },
            children: [new TextRun({ text: String(h), font: "Arial", size: 16, bold: false, color: C.mgray })]
          })]
        }))
      }),
      new TableRow({
        children: row.map((v, i) => new TableCell({
          borders: { top: { style: BorderStyle.SINGLE, size: 1, color: "D0D8E8" }, bottom: goldBorder, left: goldBorder, right: goldBorder },
          width: { size: colWidths[i], type: WidthType.DXA },
          shading: { fill: C.white, type: ShadingType.CLEAR },
          margins: { top: 120, bottom: 120, left: 120, right: 120 },
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 0, after: 0 },
            children: [new TextRun({ text: String(v ?? ""), font: "Arial", size: 22, bold: true, color: C.navy })]
          })]
        }))
      })
    ]
  });
}

// ── Spacer paragraph ──────────────────────────────────────────────
function spacer(before = 0, after = 120) {
  return new Paragraph({
    spacing: { before, after },
    children: [new TextRun({ text: "", font: "Arial", size: 2 })]
  });
}

// ── Cover section ─────────────────────────────────────────────────
function buildCoverSection(report) {
  const periodMatch = report.title.match(/(\d{4}-\d{2}-\d{2} to \d{4}-\d{2}-\d{2})/);
  const period = periodMatch ? periodMatch[1] : "";
  const parts = report.title.split("–").map(s => s.trim());
  const clientName = parts.length >= 2 ? parts[1] : report.title;

  const rule = (color, size, before, after) => new Paragraph({
    border: { bottom: { style: BorderStyle.SINGLE, size, color, space: 1 } },
    spacing: { before, after },
    children: [new TextRun({ text: "", font: "Arial", size: 2 })]
  });

  const centered = (text, size, color, bold, before, after) => new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before, after },
    children: [new TextRun({ text, font: "Arial", size, color, bold: !!bold })]
  });

  return {
    properties: {
      page: {
        size: { width: PAGE_W, height: 15840 },
        margin: { top: 2520, bottom: 2520, left: 1800, right: 1800 }
      }
    },
    children: [
      centered("CONFIDENTIAL  —  FOR INTERNAL USE ONLY", 16, C.gold, true, 0, 320),
      rule(C.gold, 24, 0, 480),
      centered("SEO  /  GEO  /  AEO  STRATEGIC REPORT", 20, C.lgray, false, 0, 200),
      centered(clientName, 52, C.navy, true, 0, 160),
      centered(report.title, 18, C.lgray, false, 0, 440),
      rule(C.gold, 24, 0, 560),
      centered("REPORTING PERIOD", 17, C.lgray, false, 0, 100),
      centered(period, 28, C.navy, true, 0, 560),
      centered("Prepared by", 17, C.lgray, false, 0, 80),
      centered("Digital Growth Strategy Team", 22, C.gold, true, 0, 0),
      spacer(640, 0),
      new Paragraph({
        border: { top: { style: BorderStyle.SINGLE, size: 6, color: C.gold, space: 1 } },
        spacing: { before: 0, after: 0 },
        children: [new TextRun({ text: "", font: "Arial", size: 2 })]
      })
    ]
  };
}

// ── Header / Footer ───────────────────────────────────────────────
function buildHeader(clientName) {
  return new Header({
    children: [
      new Paragraph({
        border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: C.gold, space: 4 } },
        spacing: { before: 0, after: 100 },
        tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
        children: [
          new TextRun({ text: clientName, font: "Arial", size: 17, bold: true, color: C.navy }),
          new TextRun({ text: "\t", font: "Arial", size: 17 }),
          new TextRun({ text: "SEO / GEO / AEO Strategic Report", font: "Arial", size: 17, color: C.lgray })
        ]
      })
    ]
  });
}

function buildFooter() {
  return new Footer({
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        border: { top: { style: BorderStyle.SINGLE, size: 4, color: "DDDDDD", space: 4 } },
        spacing: { before: 80, after: 0 },
        children: [
          new TextRun({ text: "Page ", font: "Arial", size: 17, color: C.lgray }),
          new TextRun({ children: [PageNumber.CURRENT], font: "Arial", size: 17, color: C.lgray }),
          new TextRun({ text: " of ", font: "Arial", size: 17, color: C.lgray }),
          new TextRun({ children: [PageNumber.TOTAL_PAGES], font: "Arial", size: 17, color: C.lgray })
        ]
      })
    ]
  });
}

// ── Section heading (H1) ──────────────────────────────────────────
// Always page-breaks before (cover is already a separate section,
// so the very first body H1 still gets its own page)
function buildH1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    pageBreakBefore: true,
    children: [new TextRun({ text, font: "Arial" })]
  });
}

// ── Subheading (H2) ───────────────────────────────────────────────
function buildH2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    children: [new TextRun({ text, font: "Arial" })]
  });
}

// ── Main document builder ─────────────────────────────────────────
async function buildDocx(report) {
  // Derive client name for dynamic header
  const titleParts = report.title.split("–").map(s => s.trim());
  const clientName = titleParts.length >= 2 ? titleParts[1] : "Client";

  // Filter out any COVER PAGE section from JSON (formatter handles cover)
  const contentSections = report.sections.filter(
    s => !s.heading.toUpperCase().includes("COVER")
  );

  const coverSection = buildCoverSection(report);
  const bodyChildren = [];

  for (const section of contentSections) {
    // H1 — always page breaks before
    bodyChildren.push(buildH1(section.heading));

    for (const sub of section.subsections) {
      bodyChildren.push(buildH2(sub.subheading));

      if (sub.type === "paragraph") {
        bodyChildren.push(new Paragraph({
          spacing: { before: 40, after: 140, line: 288, lineRule: "auto" },
          children: [new TextRun({ text: sub.content, font: "Arial", size: 20, color: C.body })]
        }));

      } else if (sub.type === "bullets") {
        for (const item of sub.items) {
          bodyChildren.push(new Paragraph({
            numbering: { reference: "bullets", level: 0 },
            spacing: { before: 40, after: 60 },
            children: [new TextRun({ text: String(item), font: "Arial", size: 19, color: C.body })]
          }));
        }
        bodyChildren.push(spacer(0, 80));

      } else if (sub.type === "table") {
        bodyChildren.push(spacer(40, 60));
        bodyChildren.push(buildTable(sub.headers, sub.rows));
        bodyChildren.push(spacer(0, 160));

      } else if (sub.type === "callout") {
        // Single-row highlight box for commercial estimates etc.
        bodyChildren.push(spacer(40, 60));
        bodyChildren.push(buildCallout(sub.headers, sub.row));
        bodyChildren.push(spacer(0, 160));
      }
    }
  }

  const bodySection = {
    properties: {
      page: {
        size: { width: PAGE_W, height: 15840 },
        margin: { top: 1080, bottom: 1080, left: MARGIN, right: MARGIN }
      }
    },
    headers: { default: buildHeader(clientName) },
    footers: { default: buildFooter() },
    children: bodyChildren
  };

  const doc = new Document({
    numbering: {
      config: [
        {
          reference: "bullets",
          levels: [{
            level: 0,
            format: LevelFormat.BULLET,
            text: "\u25A0",
            alignment: AlignmentType.LEFT,
            style: {
              run: { color: C.gold, size: 16, font: "Arial" },
              paragraph: { indent: { left: 560, hanging: 320 } }
            }
          }]
        }
      ]
    },
    styles: {
      default: {
        document: { run: { font: "Arial", size: 20, color: C.body } }
      },
      paragraphStyles: [
        {
          id: "Heading1", name: "Heading 1",
          basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: 30, bold: true, color: C.navy, font: "Arial" },
          paragraph: {
            spacing: { before: 0, after: 180 },  // no top space — page break handles it
            outlineLevel: 0,
            border: {
              left:   { style: BorderStyle.SINGLE, size: 18, color: C.gold, space: 8 },
              bottom: { style: BorderStyle.SINGLE, size: 4,  color: C.gold, space: 4 }
            }
          }
        },
        {
          id: "Heading2", name: "Heading 2",
          basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: 22, bold: true, color: C.blue, font: "Arial" },
          paragraph: {
            spacing: { before: 220, after: 80 },  // tight — content follows closely
            outlineLevel: 1,
            indent: { left: 120 }
          }
        }
      ]
    },
    sections: [coverSection, bodySection]
  });

  return await Packer.toBuffer(doc);
}

// ── Routes ─────────────────────────────────────────────────────────
app.post('/generate', async (req, res) => {
  try {
    const buffer = await buildDocx(req.body);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', 'attachment; filename="report.docx"');
    res.send(buffer);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

app.post('/generate-link', async (req, res) => {
  try {
    const buffer = await buildDocx(req.body);
    const filename = `report_${Date.now()}.docx`;
    const filepath = `/tmp/${filename}`;
    fs.writeFileSync(filepath, buffer);
    res.json({
      download_url: `https://docx-formatter-iobk.onrender.com/download/${filename}`,
      message: "Report generated successfully"
    });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

app.get('/download/:filename', (req, res) => {
  const filepath = `/tmp/${req.params.filename}`;
  if (fs.existsSync(filepath)) {
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', 'attachment; filename="SEO_Report.docx"');
    res.send(fs.readFileSync(filepath));
  } else {
    res.status(404).json({ error: 'File not found' });
  }
});

app.get('/health', (req, res) => res.send('OK'));

app.listen(process.env.PORT || 3000, () => console.log('Formatter running'));
