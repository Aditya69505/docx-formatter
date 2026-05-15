const express = require('express');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, BorderStyle, WidthType, ShadingType,
  LevelFormat, AlignmentType, PageBreak
} = require('docx');

const app = express();
app.use(express.json({ limit: '10mb' }));

const BRAND = "1A56A0", GRAY = "444444", WHITE = "FFFFFF", LIGHT = "EEF4FB";

function makeCell(text, isHeader, width) {
  const b = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
  return new TableCell({
    borders: { top: b, bottom: b, left: b, right: b },
    width: { size: width, type: WidthType.DXA },
    shading: { fill: isHeader ? BRAND : WHITE, type: ShadingType.CLEAR },
    margins: { top: 80, bottom: 80, left: 100, right: 100 },
    children: [new Paragraph({
      children: [new TextRun({
        text: String(text ?? ""), font: "Arial", size: 18,
        bold: isHeader, color: isHeader ? WHITE : GRAY
      })]
    })]
  });
}

function buildTable(headers, rows) {
  const colWidth = Math.floor(9026 / headers.length);
  return new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: Array(headers.length).fill(colWidth),
    rows: [
      new TableRow({
        tableHeader: true,
        children: headers.map(h => makeCell(h, true, colWidth))
      }),
      ...rows.map((r, i) => new TableRow({
        children: r.map(c => {
          const cell = makeCell(c, false, colWidth);
          if (i % 2 === 1) cell.options.shading = { fill: LIGHT, type: ShadingType.CLEAR };
          return cell;
        })
      }))
    ]
  });
}

app.post('/generate', async (req, res) => {
  try {
    const report = req.body;
    const children = [];

    children.push(new Paragraph({
      heading: HeadingLevel.TITLE,
      children: [new TextRun({ text: report.title, font: "Arial", bold: true, color: BRAND })]
    }));

    for (const section of report.sections) {
      children.push(new Paragraph({ children: [new PageBreak()] }));
      children.push(new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun({ text: section.heading, font: "Arial" })]
      }));

      for (const sub of section.subsections) {
        children.push(new Paragraph({
          heading: HeadingLevel.HEADING_2,
          spacing: { before: 200, after: 80 },
          children: [new TextRun({ text: sub.subheading, font: "Arial" })]
        }));

        if (sub.type === "paragraph") {
          children.push(new Paragraph({
            spacing: { before: 60, after: 80 },
            children: [new TextRun({ text: sub.content, font: "Arial", size: 20, color: GRAY })]
          }));

        } else if (sub.type === "bullets") {
          for (const item of sub.items) {
            children.push(new Paragraph({
              numbering: { reference: "bullets", level: 0 },
              spacing: { before: 40, after: 40 },
              children: [new TextRun({ text: item, font: "Arial", size: 20, color: GRAY })]
            }));
          }

        } else if (sub.type === "table") {
          children.push(buildTable(sub.headers, sub.rows));
          children.push(new Paragraph({ children: [new TextRun("")] }));
        }
      }
    }

    const doc = new Document({
      numbering: {
        config: [{
          reference: "bullets",
          levels: [{
            level: 0, format: LevelFormat.BULLET, text: "•",
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720, hanging: 360 } } }
          }]
        }]
      },
      styles: {
        default: { document: { run: { font: "Arial", size: 22 } } },
        paragraphStyles: [
          { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal",
            run: { size: 36, bold: true, color: BRAND },
            paragraph: { spacing: { before: 400, after: 200 }, outlineLevel: 0 } },
          { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal",
            run: { size: 26, bold: true, color: "1A3A6A" },
            paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 1 } }
        ]
      },
      sections: [{ children }]
    });

    const buffer = await Packer.toBuffer(doc);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', 'attachment; filename="report.docx"');
    res.send(buffer);

  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/generate-link', async (req, res) => {
  try {
    // ... same docx generation code ...
    const buffer = await Packer.toBuffer(doc);
    const filename = `report_${Date.now()}.docx`;
    const filepath = `/tmp/${filename}`;
    fs.writeFileSync(filepath, buffer);
    
    // Return the download URL
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
    res.setHeader('Content-Disposition', `attachment; filename="SEO_Report.docx"`);
    res.send(fs.readFileSync(filepath));
  } else {
    res.status(404).json({ error: 'File not found' });
  }
});

app.get('/health', (req, res) => res.send('OK'));
app.listen(process.env.PORT || 3000, () => console.log('Formatter running'));
