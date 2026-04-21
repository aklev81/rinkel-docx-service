const express = require('express');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, BorderStyle, WidthType, ShadingType,
  LevelFormat, PageNumber, Footer, Header, TabStopType
} = require('docx');

const app = express();
app.use(express.json({ limit: '10mb' }));

// ── HUISSTIJL KLEUREN ──────────────────────────────────────────────────────────
const BLUE      = "1A4F8A";
const MID_BLUE  = "2E75B6";
const LIGHT_BLUE = "D6E4F0";
const ACCENT    = "EAF3FB";
const GRAY      = "666666";
const BORDER_COLOR = "CCCCCC";
const WHITE     = "FFFFFF";
const GREEN     = "27AE60";
const ORANGE    = "E67E22";
const RED       = "C0392B";

// ── BORDER HELPERS ─────────────────────────────────────────────────────────────
const border    = { style: BorderStyle.SINGLE, size: 1, color: BORDER_COLOR };
const borders   = { top: border, bottom: border, left: border, right: border };
const noBorder  = { style: BorderStyle.NONE, size: 0, color: WHITE };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

// ── BUILDER HELPERS ────────────────────────────────────────────────────────────
function spacer(size = 120) {
  return new Paragraph({ spacing: { before: 0, after: size }, children: [new TextRun("")] });
}

function heading1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 400, after: 160 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: MID_BLUE, space: 4 } },
    children: [new TextRun({ text, bold: true, size: 32, color: BLUE, font: "Arial" })]
  });
}

function heading2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 280, after: 120 },
    children: [new TextRun({ text, bold: true, size: 26, color: MID_BLUE, font: "Arial" })]
  });
}

function heading3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    spacing: { before: 200, after: 80 },
    children: [new TextRun({ text, bold: true, size: 22, color: "2C3E50", font: "Arial" })]
  });
}

function body(text, options = {}) {
  return new Paragraph({
    spacing: { before: 60, after: 100 },
    children: [new TextRun({ text, size: 22, font: "Arial", ...options })]
  });
}

function bullet(text) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { before: 40, after: 60 },
    children: [new TextRun({ text, size: 22, font: "Arial" })]
  });
}

function note(text) {
  return new Paragraph({
    spacing: { before: 40, after: 80 },
    children: [new TextRun({ text, size: 20, font: "Arial", color: GRAY, italics: true })]
  });
}

function infoBox(label, content) {
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [1800, 7560],
    rows: [new TableRow({
      children: [
        new TableCell({
          borders, width: { size: 1800, type: WidthType.DXA },
          shading: { fill: LIGHT_BLUE, type: ShadingType.CLEAR },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          children: [new Paragraph({ children: [new TextRun({ text: label, bold: true, size: 20, font: "Arial", color: BLUE })] })]
        }),
        new TableCell({
          borders, width: { size: 7560, type: WidthType.DXA },
          shading: { fill: WHITE, type: ShadingType.CLEAR },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          children: [new Paragraph({ children: [new TextRun({ text: content, size: 20, font: "Arial" })] })]
        })
      ]
    })]
  });
}

function plgBox(text) {
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [9360],
    rows: [new TableRow({
      children: [new TableCell({
        borders,
        width: { size: 9360, type: WidthType.DXA },
        shading: { fill: ACCENT, type: ShadingType.CLEAR },
        margins: { top: 160, bottom: 160, left: 200, right: 200 },
        children: [
          new Paragraph({ spacing: { before: 0, after: 60 }, children: [new TextRun({ text: "Over dit rapport", bold: true, size: 22, font: "Arial", color: BLUE })] }),
          new Paragraph({ children: [new TextRun({ text, size: 20, font: "Arial", color: "2C3E50" })] })
        ]
      })]
    })]
  });
}

function visionBox(text) {
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [9360],
    rows: [new TableRow({
      children: [new TableCell({
        borders,
        width: { size: 9360, type: WidthType.DXA },
        shading: { fill: ACCENT, type: ShadingType.CLEAR },
        margins: { top: 160, bottom: 160, left: 200, right: 200 },
        children: [
          new Paragraph({ spacing: { before: 0, after: 60 }, children: [new TextRun({ text: "Productvisie", bold: true, size: 22, font: "Arial", color: BLUE })] }),
          new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text, size: 22, font: "Arial", color: BLUE, bold: true, italics: true })] })
        ]
      })]
    })]
  });
}

// ── TABEL BUILDERS ─────────────────────────────────────────────────────────────

// Generieke tabel met header rij
function makeTable(headers, rows, colWidths) {
  const total = colWidths.reduce((a, b) => a + b, 0);
  const headerRow = new TableRow({
    children: headers.map((h, i) => new TableCell({
      borders,
      width: { size: colWidths[i], type: WidthType.DXA },
      shading: { fill: BLUE, type: ShadingType.CLEAR },
      margins: { top: 80, bottom: 80, left: 120, right: 120 },
      children: [new Paragraph({ children: [new TextRun({ text: h, bold: true, size: 20, font: "Arial", color: WHITE })] })]
    }))
  });

  const dataRows = rows.map((row, ri) =>
    new TableRow({
      children: row.map((cell, ci) => new TableCell({
        borders,
        width: { size: colWidths[ci], type: WidthType.DXA },
        shading: { fill: ri % 2 === 0 ? "F5F9FD" : WHITE, type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        children: [new Paragraph({ children: [new TextRun({ text: String(cell || ''), size: 20, font: "Arial" })] })]
      }))
    })
  );

  return new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: colWidths, rows: [headerRow, ...dataRows] });
}

// Actiepuntentabel (omschrijving, aandachtsgebied, prioriteit)
function actionTable(rows) {
  const priorityColor = (p) => {
    if (p && p.toLowerCase().includes('hoog')) return { fill: "FDECEA", color: RED };
    if (p && p.toLowerCase().includes('middel')) return { fill: "FFF3CD", color: ORANGE };
    return { fill: "EAF6EA", color: GREEN };
  };

  const headerRow = new TableRow({
    children: [
      new TableCell({ borders, width: { size: 400, type: WidthType.DXA }, shading: { fill: BLUE, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 100, right: 100 }, children: [new Paragraph({ children: [new TextRun({ text: "#", bold: true, size: 20, font: "Arial", color: WHITE })] })] }),
      new TableCell({ borders, width: { size: 5000, type: WidthType.DXA }, shading: { fill: BLUE, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "Actiepunt", bold: true, size: 20, font: "Arial", color: WHITE })] })] }),
      new TableCell({ borders, width: { size: 2160, type: WidthType.DXA }, shading: { fill: BLUE, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "Aandachtsgebied", bold: true, size: 20, font: "Arial", color: WHITE })] })] }),
      new TableCell({ borders, width: { size: 1800, type: WidthType.DXA }, shading: { fill: BLUE, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "Prioriteit", bold: true, size: 20, font: "Arial", color: WHITE })] })] })
    ]
  });

  const dataRows = rows.map((r, i) => {
    const pc = priorityColor(r[2]);
    return new TableRow({
      children: [
        new TableCell({ borders, width: { size: 400, type: WidthType.DXA }, shading: { fill: i % 2 === 0 ? "F5F9FD" : WHITE, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 100, right: 100 }, children: [new Paragraph({ children: [new TextRun({ text: String(i + 1), bold: true, size: 20, font: "Arial", color: MID_BLUE })] })] }),
        new TableCell({ borders, width: { size: 5000, type: WidthType.DXA }, shading: { fill: i % 2 === 0 ? "F5F9FD" : WHITE, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: r[0] || '', size: 20, font: "Arial" })] })] }),
        new TableCell({ borders, width: { size: 2160, type: WidthType.DXA }, shading: { fill: i % 2 === 0 ? "F5F9FD" : WHITE, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: r[1] || '', size: 20, font: "Arial", color: GRAY })] })] }),
        new TableCell({ borders, width: { size: 1800, type: WidthType.DXA }, shading: { fill: pc.fill, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: r[2] || '', bold: true, size: 20, font: "Arial", color: pc.color })] })] })
      ]
    });
  });

  return new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [400, 5000, 2160, 1800], rows: [headerRow, ...dataRows] });
}

// ── MARKDOWN PARSER ────────────────────────────────────────────────────────────
// Zet Claude's markdown om naar docx-elementen met Rinkel huisstijl
function parseMarkdown(text, onderzoekNaam, datum) {
  const lines = text.split('\n');
  const elements = [];
  let i = 0;
  let inTable = false;
  let tableLines = [];
  let tableSection = '';

  // Detecteer huidige sectie voor contextuele opmaak
  let currentSection = '';

  const flushTable = () => {
    if (tableLines.length < 2) { tableLines = []; inTable = false; return; }

    // Parse markdown tabel
    const headerLine = tableLines[0];
    const dataLines = tableLines.slice(2); // sla separator over

    const parseRow = (line) => line.split('|').map(c => c.trim().replace(/\*\*/g, '')).filter((c, i, arr) => i > 0 && i < arr.length - 1);

    const headers = parseRow(headerLine);
    const rows = dataLines.map(parseRow).filter(r => r.length > 0);

    if (headers.length === 0) { tableLines = []; inTable = false; return; }

    // Bereken kolombreedtes op basis van aantal kolommen
    const total = 9360;
    const colWidths = headers.map(() => Math.floor(total / headers.length));

    // Detecteer actiepuntentabel (heeft prioriteit kolom)
    const isActionTable = headers.some(h => h.toLowerCase().includes('prioriteit'));

    // Detecteer kernbevindingentabel (heeft A/B kolommen)
    const isABTable = headers.some(h => h.includes('(A)') || h.includes('(B)'));

    if (isActionTable) {
      const actionRows = rows.map(r => [r[1] || r[0], r[2] || '', r[3] || r[2] || '']);
      elements.push(actionTable(actionRows));
    } else if (isABTable) {
      // Drie-kolommentabel met brede B kolom
      const abWidths = [2000, 3000, 4360];
      elements.push(makeTable(headers, rows, abWidths));
    } else if (headers.length === 2) {
      // Informatietabel (methodologie)
      rows.forEach(r => {
        elements.push(infoBox(r[0] || '', r[1] || ''));
        elements.push(spacer(40));
      });
      tableLines = []; inTable = false; return;
    } else {
      // Standaard tabel
      const evenWidths = headers.map(() => Math.floor(total / headers.length));
      elements.push(makeTable(headers, rows, evenWidths));
    }

    elements.push(spacer(80));
    tableLines = []; inTable = false;
  };

  while (i < lines.length) {
    const line = lines[i];
    const trimmed = line.trim();

    // Tabel detectie
    if (trimmed.startsWith('|')) {
      if (!inTable) inTable = true;
      tableLines.push(trimmed);
      i++; continue;
    } else if (inTable) {
      flushTable();
    }

    // Lege regel
    if (trimmed === '' || trimmed === '---') {
      elements.push(spacer(60));
      i++; continue;
    }

    // Headers
    if (trimmed.startsWith('#### ')) {
      elements.push(heading3(trimmed.replace(/^####\s+/, '').replace(/\*\*/g, '')));
      i++; continue;
    }
    if (trimmed.startsWith('### ')) {
      const text = trimmed.replace(/^###\s+/, '').replace(/\*\*/g, '');
      elements.push(heading3(text));
      i++; continue;
    }
    if (trimmed.startsWith('## ')) {
      const text = trimmed.replace(/^##\s+/, '').replace(/\*\*/g, '');
      currentSection = text;
      elements.push(heading2(text));
      i++; continue;
    }
    if (trimmed.startsWith('# ')) {
      const text = trimmed.replace(/^#\s+/, '').replace(/\*\*/g, '');
      currentSection = text;
      elements.push(heading1(text));
      i++; continue;
    }

    // Vetgedrukte headers die als sectietitels worden gebruikt (bijv. **3.1 Titel**)
    const boldHeader = trimmed.match(/^\*\*([0-9]+\.[0-9]*\s+.+?)\*\*$/) ||
                       trimmed.match(/^\*\*(BIJLAGE\s+.+?)\*\*$/) ||
                       trimmed.match(/^\*\*(PLG-INLEIDING.*?)\*\*$/);
    if (boldHeader) {
      const text = boldHeader[1];
      // Bepaal level op basis van nummering
      if (text.match(/^[0-9]+\.\s/)) {
        elements.push(heading1(text));
      } else if (text.match(/^[0-9]+\.[0-9]+\s/)) {
        elements.push(heading2(text));
      } else {
        elements.push(heading2(text));
      }
      i++; continue;
    }

    // Bullets
    if (trimmed.startsWith('- ') || trimmed.startsWith('• ')) {
      const text = trimmed.replace(/^[-•]\s+/, '').replace(/\*\*/g, '');
      elements.push(bullet(text));
      i++; continue;
    }

    // Genummerde lijsten
    if (trimmed.match(/^[0-9]+\.\s/)) {
      const text = trimmed.replace(/^[0-9]+\.\s+/, '').replace(/\*\*/g, '');
      elements.push(bullet(text));
      i++; continue;
    }

    // Noten (cursief of grijs)
    if (trimmed.startsWith('*Noot') || trimmed.startsWith('_Noot') || trimmed.startsWith('> ')) {
      const text = trimmed.replace(/^[*_>]\s*/, '').replace(/[*_]$/, '');
      elements.push(note(text));
      i++; continue;
    }

    // Productvisie kader
    if (trimmed.toLowerCase().includes('productvisie') && trimmed.startsWith('*') && trimmed.endsWith('*')) {
      const text = trimmed.replace(/^\*+/, '').replace(/\*+$/, '');
      elements.push(visionBox(text));
      i++; continue;
    }

    // Geciteerde zin (productvisie in aanhalingstekens)
    if ((trimmed.startsWith('"') && trimmed.endsWith('"')) ||
        (trimmed.startsWith('\"') && trimmed.endsWith('\"'))) {
      const text = trimmed.replace(/^[""]/, '').replace(/[""]$/, '');
      elements.push(visionBox(text));
      i++; continue;
    }

    // PLG inleiding kader
    if (trimmed.toLowerCase().startsWith('dit rapport is geschreven vanuit een product led')) {
      elements.push(plgBox(trimmed.replace(/\*\*/g, '')));
      i++; continue;
    }

    // Gewone alinea — strip markdown opmaak
    if (trimmed.length > 0) {
      const clean = trimmed
        .replace(/\*\*/g, '')
        .replace(/\*/g, '')
        .replace(/_{2}/g, '')
        .replace(/_/g, '')
        .replace(/\[AANNAME\]/g, '[AANNAME]');
      elements.push(body(clean));
    }

    i++;
  }

  // Flush eventuele resterende tabel
  if (inTable && tableLines.length > 0) flushTable();

  return elements;
}

// ── DOCUMENT BUILDER ───────────────────────────────────────────────────────────
async function buildDocument(rapportTekst, onderzoekNaam, datum) {
  const contentElements = parseMarkdown(rapportTekst, onderzoekNaam, datum);

  const doc = new Document({
    numbering: {
      config: [{
        reference: "bullets",
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 600, hanging: 300 } } }
        }]
      }]
    },
    styles: {
      default: { document: { run: { font: "Arial", size: 22 } } },
      paragraphStyles: [
        { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 32, bold: true, font: "Arial", color: BLUE }, paragraph: { spacing: { before: 400, after: 160 }, outlineLevel: 0 } },
        { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 26, bold: true, font: "Arial", color: MID_BLUE }, paragraph: { spacing: { before: 280, after: 120 }, outlineLevel: 1 } },
        { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 22, bold: true, font: "Arial", color: "2C3E50" }, paragraph: { spacing: { before: 200, after: 80 }, outlineLevel: 2 } }
      ]
    },
    sections: [{
      properties: {
        page: { size: { width: 11906, height: 16838 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } }
      },
      headers: {
        default: new Header({
          children: [new Paragraph({
            border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: MID_BLUE, space: 4 } },
            children: [
              new TextRun({ text: `Rinkel.com — ${onderzoekNaam}`, size: 18, font: "Arial", color: GRAY }),
              new TextRun({ text: "    |    Vertrouwelijk", size: 18, font: "Arial", color: BORDER_COLOR })
            ]
          })]
        })
      },
      footers: {
        default: new Footer({
          children: [new Paragraph({
            border: { top: { style: BorderStyle.SINGLE, size: 4, color: BORDER_COLOR, space: 4 } },
            children: [
              new TextRun({ text: `© Rinkel.com ${new Date().getFullYear()}  |  Intern gebruik  |  Pagina `, size: 18, font: "Arial", color: GRAY }),
              new TextRun({ children: [PageNumber.CURRENT], size: 18, font: "Arial", color: GRAY }),
              new TextRun({ text: " van ", size: 18, font: "Arial", color: GRAY }),
              new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 18, font: "Arial", color: GRAY })
            ]
          })]
        })
      },
      children: [
        // Coverpagina
        new Paragraph({ spacing: { before: 800, after: 200 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: "KLANTONDERZOEK — PRODUCT LED PERSPECTIEF", size: 20, font: "Arial", color: MID_BLUE, bold: true, allCaps: true })] }),
        new Paragraph({ spacing: { before: 0, after: 120 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Wat moet het product doen om zichzelf te verkopen?", size: 44, font: "Arial", color: BLUE, bold: true })] }),
        new Paragraph({ spacing: { before: 0, after: 80 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: onderzoekNaam, size: 28, font: "Arial", color: MID_BLUE, bold: true })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 400 }, children: [new TextRun({ text: `Rinkel.com  |  ${datum}  |  Vertrouwelijk`, size: 20, font: "Arial", color: GRAY })] }),
        new Paragraph({ border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: LIGHT_BLUE, space: 1 } }, children: [new TextRun("")] }),
        spacer(200),

        // Rapport inhoud
        ...contentElements,

        // Afsluiting
        spacer(120),
        body("— Einde rapport —", { color: GRAY, italics: true }),
      ]
    }]
  });

  return await Packer.toBuffer(doc);
}

// ── ROUTES ─────────────────────────────────────────────────────────────────────
app.post('/generate', async (req, res) => {
  try {
    const { rapportTekst, onderzoekNaam = 'Rinkel Onderzoek', datum = new Date().toLocaleDateString('nl-NL') } = req.body;

    if (!rapportTekst) {
      return res.status(400).json({ error: 'rapportTekst is verplicht' });
    }

    const buffer = await buildDocument(rapportTekst, onderzoekNaam, datum);

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="${onderzoekNaam}.docx"`);
    res.send(buffer);

  } catch (err) {
    console.error('Fout bij genereren rapport:', err);
    res.status(500).json({ error: err.message });
  }
});

app.get('/health', (req, res) => res.json({ status: 'ok' }));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Rinkel rapport server draait op poort ${PORT}`));
