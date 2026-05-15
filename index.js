const express = require('express');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, BorderStyle, WidthType, ShadingType, LevelFormat,
  AlignmentType, PageBreak, Header, Footer, PageNumber,
  TabStopType, TabStopPosition
} = require('docx');
const fs = require('fs');

const app = express();
app.use(express.json({ limit: '10mb' }));

// ── Color Palette ──────────────────────────────────────────────────
const NAVY     = "1A2B4A";
const GOLD     = "C8922A";
const BLUE     = "2E5090";
const BODY     = "333333";
const LGRAY    = "888888";
const ALTROW   = "F0F4FA";
const WHITE    = "FFFFFF";

// ── Priority & Status Colors ───────────────────────────────────────
function priorityColor(val) {
  const v = String(val).trim();
  if (v === "P0") return "C0392B";
  if (v === "P1") return "E67E22";
  if (v === "P2") return "27AE60";
  return null;
}

function statusColor(val) {
  const v = String(val);
  if (v.includes("Verified Data"))        return "1A7A3A";
  if (v.includes("Calculated Estimate"))  return "1A5EA0";
  if (v.includes("Strategic Assumption")) return "8B6914";
  if (v.includes("Needs Verification"))   return "A03020";
  return null;
}

// ── Table Builder ──────────────────────────────────────────────────
function makeCell(text, isHeader, width, rowIdx) {
  const stdBorder  = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
  const goldBorder = { style: BorderStyle.SINGLE, size: 3, color: GOLD };

  const textStr = String(text ?? "");
  const pc = !isHeader ? priorityColor(textStr) : null;
  const sc = !isHeader ? statusColor(textStr) : null;
  const textColor = isHeader ? WHITE : (pc || sc || BODY);
  const isBold = isHeader || !!pc;

  const fill = isHeader ? NAVY : (rowIdx % 2 === 1 ? ALTROW : WHITE);

  return new TableCell({
    borders: {
      top:    isHeader ? goldBorder : stdBorder,
      bottom: isHeader ? goldBorder : stdBorder,
      left:   stdBorder,
      right:  stdBorder
    },
    width: { size: width, type: WidthType.DXA },
    shading: { fill, type: ShadingType.CLEAR },
    margins: { top: 100, bottom: 100, left: 140, right: 140 },
    children: [new Paragraph({
      alignment: isHeader ? AlignmentType.CENTER : AlignmentType.LEFT,
      children: [new TextRun({
        text: textStr,
        font: "Arial",
        size: 18,
        bold: isBold,
        color: textColor
      })]
    })]
  });
}

function buildTable(headers, rows) {
  const TABLE_W  = 9720;
  const colCount = headers.length;

  let colWidths;
  if (colCount === 1) {
    colWidths = [TABLE_W];
  } else {
    const firstColW = Math.floor(TABLE_W * 0.22);
    const otherColW = Math.floor((TABLE_W - firstColW) / (colCount - 1));
    colWidths = [firstColW, ...Array(colCount - 1).fill(otherColW)];
    const diff = TABLE_W - colWidths.reduce((a, b) => a + b, 0);
    colWidths[colWidths.length - 1] += diff;
  }

  return new Table({
    width: { size: TABLE_W, type: WidthType.DXA },
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

// ── Cover Page Section ─────────────────────────────────────────────
function buildCoverSection(report) {
  // Extract period from title e.g. "... – 2026-04-08 to 2026-05-08"
  const periodMatch = report.title.match(/(\d{4}-\d{2}-\d{2} to \d{4}-\d{2}-\d{2})/);
  const period = periodMatch ? periodMatch[1] : "";
  // Extract client name — text between first and second "–"
  const parts = report.title.split("–").map(s => s.trim());
  const clientName = parts.length >= 2 ? parts[1] : report.title;

  return {
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 2520, bottom: 2520, left: 1800, right: 1800 }
      }
    },
    children: [
      // Confidential label
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 360 },
        children: [new TextRun({
          text: "CONFIDENTIAL  —  FOR INTERNAL USE ONLY",
          font: "Arial", size: 16, color: GOLD, bold: true
        })]
      }),

      // Gold rule
      new Paragraph({
        alignment: AlignmentType.CENTER,
        border: { bottom: { style: BorderStyle.SINGLE, size: 24, color: GOLD, space: 1 } },
        spacing: { before: 0, after: 560 },
        children: [new TextRun({ text: "", font: "Arial", size: 2 })]
      }),

      // Report type
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 240 },
        children: [new TextRun({
          text: "SEO  /  GEO  /  AEO  STRATEGIC REPORT",
          font: "Arial", size: 22, color: LGRAY
        })]
      }),

      // Client name (big gold)
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 200 },
        children: [new TextRun({
          text: clientName,
          font: "Arial", size: 56, bold: true, color: NAVY
        })]
      }),

      // Full title line (smaller navy)
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 480 },
        children: [new TextRun({
          text: report.title,
          font: "Arial", size: 20, color: LGRAY
        })]
      }),

      // Gold rule below title
      new Paragraph({
        alignment: AlignmentType.CENTER,
        border: { bottom: { style: BorderStyle.SINGLE, size: 24, color: GOLD, space: 1 } },
        spacing: { before: 0, after: 640 },
        children: [new TextRun({ text: "", font: "Arial", size: 2 })]
      }),

      // Reporting Period label
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 120 },
        children: [new TextRun({
          text: "REPORTING PERIOD",
          font: "Arial", size: 18, color: LGRAY, bold: false
        })]
      }),

      // Period value
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 640 },
        children: [new TextRun({
          text: period,
          font: "Arial", size: 30, color: NAVY, bold: true
        })]
      }),

      // Prepared by label
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 100 },
        children: [new TextRun({
          text: "Prepared by",
          font: "Arial", size: 18, color: LGRAY
        })]
      }),

      // Team name
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 800 },
        children: [new TextRun({
          text: "Digital Growth Strategy Team",
          font: "Arial", size: 24, color: GOLD, bold: true
        })]
      }),

      // Bottom thin gold rule
      new Paragraph({
        alignment: AlignmentType.CENTER,
        border: { top: { style: BorderStyle.SINGLE, size: 6, color: GOLD, space: 1 } },
        spacing: { before: 0, after: 0 },
        children: [new TextRun({ text: "", font: "Arial", size: 2 })]
      })
    ]
  };
}

// ── Main buildDocx ─────────────────────────────────────────────────
async function buildDocx(report) {
  const coverSection = buildCoverSection(report);

  const bodyChildren = [];
  let isFirstSection = true;

  for (const section of report.sections) {
    // H1 — with page break on all except first
    bodyChildren.push(new Paragraph({
      heading: HeadingLevel.HEADING_1,
      pageBreakBefore: !isFirstSection,
      children: [new TextRun({ text: section.heading, font: "Arial" })]
    }));
    isFirstSection = false;

    for (const sub of section.subsections) {
      // Subtle divider before each H2
      bodyChildren.push(new Paragraph({
        spacing: { before: 160, after: 160 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 1, color: "E0E4EC", space: 1 } },
        children: [new TextRun({ text: "", size: 2 })]
      }));

      // H2
      bodyChildren.push(new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun({ text: sub.subheading, font: "Arial" })]
      }));

      if (sub.type === "paragraph") {
        bodyChildren.push(new Paragraph({
          spacing: { before: 60, after: 100, line: 276, lineRule: "auto" },
          children: [new TextRun({ text: sub.content, font: "Arial", size: 20, color: BODY })]
        }));

      } else if (sub.type === "bullets") {
        for (const item of sub.items) {
          bodyChildren.push(new Paragraph({
            numbering: { reference: "bullets", level: 0 },
            spacing: { before: 60, after: 60 },
            children: [new TextRun({ text: item, font: "Arial", size: 20, color: BODY })]
          }));
        }

      } else if (sub.type === "table") {
        bodyChildren.push(buildTable(sub.headers, sub.rows));
        bodyChildren.push(new Paragraph({
          spacing: { before: 0, after: 200 },
          children: [new TextRun({ text: "", font: "Arial", size: 4 })]
        }));
      }
    }
  }

  const bodySection = {
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1080, bottom: 1080, left: 1260, right: 1260 }
      }
    },
    headers: {
      default: new Header({
        children: [
          new Paragraph({
            border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: GOLD, space: 4 } },
            spacing: { before: 0, after: 120 },
            tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
            children: [
              new TextRun({ text: "Crunchy Fashion", font: "Arial", size: 18, bold: true, color: NAVY }),
              new TextRun({ text: "\t", font: "Arial", size: 18 }),
              new TextRun({ text: "SEO / GEO / AEO Strategic Report", font: "Arial", size: 18, color: LGRAY })
            ]
          })
        ]
      })
    },
    footers: {
      default: new Footer({
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            border: { top: { style: BorderStyle.SINGLE, size: 4, color: "DDDDDD", space: 4 } },
            spacing: { before: 80, after: 0 },
            children: [
              new TextRun({ text: "Page ", font: "Arial", size: 18, color: LGRAY }),
              new TextRun({ children: [PageNumber.CURRENT], font: "Arial", size: 18, color: LGRAY }),
              new TextRun({ text: " of ", font: "Arial", size: 18, color: LGRAY }),
              new TextRun({ children: [PageNumber.TOTAL_PAGES], font: "Arial", size: 18, color: LGRAY })
            ]
          })
        ]
      })
    },
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
              run: { color: GOLD, size: 16, font: "Arial" },
              paragraph: { indent: { left: 720, hanging: 360 } }
            }
          }]
        },
        {
          reference: "numbers",
          levels: [{
            level: 0,
            format: LevelFormat.DECIMAL,
            text: "%1.",
            alignment: AlignmentType.LEFT,
            style: {
              paragraph: { indent: { left: 720, hanging: 360 } }
            }
          }]
        }
      ]
    },
    styles: {
      default: {
        document: { run: { font: "Arial", size: 20, color: BODY } }
      },
      paragraphStyles: [
        {
          id: "Heading1", name: "Heading 1",
          basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: 32, bold: true, color: NAVY, font: "Arial" },
          paragraph: {
            spacing: { before: 480, after: 200 },
            outlineLevel: 0,
            border: {
              left:   { style: BorderStyle.SINGLE, size: 20, color: GOLD, space: 8 },
              bottom: { style: BorderStyle.SINGLE, size: 6,  color: GOLD, space: 4 }
            }
          }
        },
        {
          id: "Heading2", name: "Heading 2",
          basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: 24, bold: true, color: BLUE, font: "Arial" },
          paragraph: {
            spacing: { before: 280, after: 120 },
            outlineLevel: 1,
            indent: { left: 160 }
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
