const express = require('express');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, BorderStyle, WidthType, ShadingType,
  LevelFormat, PageNumber, Footer, Header
} = require('docx');

const app = express();
app.use(express.json({ limit: '10mb' }));

// ── KLEUREN ────────────────────────────────────────────────────────────────────
const BLUE        = "1A4F8A";
const LIGHT_BLUE  = "D6E4F0";
const MID_BLUE    = "2E75B6";
const GRAY        = "666666";
const BORDER_COLOR = "CCCCCC";

const border   = { style: BorderStyle.SINGLE, size: 1, color: BORDER_COLOR };
const borders  = { top: border, bottom: border, left: border, right: border };

// ── HELPER FUNCTIES ────────────────────────────────────────────────────────────
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

function spacer(size = 120) {
  return new Paragraph({
    spacing: { before: 0, after: size },
    children: [new TextRun("")]
  });
}

function infoBox(label, content) {
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [1500, 7860],
    rows: [
      new TableRow({
        children: [
          new TableCell({
            borders,
            width: { size: 1500, type: WidthType.DXA },
            shading: { fill: LIGHT_BLUE, type: ShadingType.CLEAR },
            margins: { top: 100, bottom: 100, left: 140, right: 140 },
            verticalAlign: "center",
            children: [new Paragraph({ children: [new TextRun({ text: label, bold: true, size: 20, font: "Arial", color: BLUE })] })]
          }),
          new TableCell({
            borders,
            width: { size: 7860, type: WidthType.DXA },
            shading: { fill: "FFFFFF", type: ShadingType.CLEAR },
            margins: { top: 100, bottom: 100, left: 140, right: 140 },
            children: [new Paragraph({ children: [new TextRun({ text: content, size: 20, font: "Arial" })] })]
          })
        ]
      })
    ]
  });
}

function metricTable(rows) {
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [4680, 4680],
    rows: [
      new TableRow({
        children: [
          new TableCell({
            borders,
            width: { size: 4680, type: WidthType.DXA },
            shading: { fill: BLUE, type: ShadingType.CLEAR },
            margins: { top: 100, bottom: 100, left: 140, right: 140 },
            children: [new Paragraph({ children: [new TextRun({ text: "Bevinding", bold: true, size: 22, font: "Arial", color: "FFFFFF" })] })]
          }),
          new TableCell({
            borders,
            width: { size: 4680, type: WidthType.DXA },
            shading: { fill: BLUE, type: ShadingType.CLEAR },
            margins: { top: 100, bottom: 100, left: 140, right: 140 },
            children: [new Paragraph({ children: [new TextRun({ text: "Detail", bold: true, size: 22, font: "Arial", color: "FFFFFF" })] })]
          })
        ]
      }),
      ...rows.map((r, i) => new TableRow({
        children: [
          new TableCell({
            borders,
            width: { size: 4680, type: WidthType.DXA },
            shading: { fill: i % 2 === 0 ? "F5F9FD" : "FFFFFF", type: ShadingType.CLEAR },
            margins: { top: 80, bottom: 80, left: 140, right: 140 },
            children: [new Paragraph({ children: [new TextRun({ text: r[0], bold: true, size: 20, font: "Arial" })] })]
          }),
          new TableCell({
            borders,
            width: { size: 4680, type: WidthType.DXA },
            shading: { fill: i % 2 === 0 ? "F5F9FD" : "FFFFFF", type: ShadingType.CLEAR },
            margins: { top: 80, bottom: 80, left: 140, right: 140 },
            children: [new Paragraph({ children: [new TextRun({ text: r[1], size: 20, font: "Arial" })] })]
          })
        ]
      }))
    ]
  });
}

function actionTable(rows) {
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [400, 4000, 2560, 2400],
    rows: [
      new TableRow({
        children: [
          new TableCell({ borders, width: { size: 400, type: WidthType.DXA }, shading: { fill: BLUE, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 100, right: 100 }, children: [new Paragraph({ children: [new TextRun({ text: "#", bold: true, size: 20, font: "Arial", color: "FFFFFF" })] })] }),
          new TableCell({ borders, width: { size: 4000, type: WidthType.DXA }, shading: { fill: BLUE, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "Actiepunt", bold: true, size: 20, font: "Arial", color: "FFFFFF" })] })] }),
          new TableCell({ borders, width: { size: 2560, type: WidthType.DXA }, shading: { fill: BLUE, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "Aandachtsgebied", bold: true, size: 20, font: "Arial", color: "FFFFFF" })] })] }),
          new TableCell({ borders, width: { size: 2400, type: WidthType.DXA }, shading: { fill: BLUE, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "Prioriteit", bold: true, size: 20, font: "Arial", color: "FFFFFF" })] })] })
        ]
      }),
      ...rows.map((r, i) => new TableRow({
        children: [
          new TableCell({ borders, width: { size: 400, type: WidthType.DXA }, shading: { fill: i % 2 === 0 ? "F5F9FD" : "FFFFFF", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 100, right: 100 }, children: [new Paragraph({ children: [new TextRun({ text: String(i + 1), bold: true, size: 20, font: "Arial", color: MID_BLUE })] })] }),
          new TableCell({ borders, width: { size: 4000, type: WidthType.DXA }, shading: { fill: i % 2 === 0 ? "F5F9FD" : "FFFFFF", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: r[0], size: 20, font: "Arial" })] })] }),
          new TableCell({ borders, width: { size: 2560, type: WidthType.DXA }, shading: { fill: i % 2 === 0 ? "F5F9FD" : "FFFFFF", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: r[1], size: 20, font: "Arial", color: GRAY })] })] }),
          new TableCell({
            borders,
            width: { size: 2400, type: WidthType.DXA },
            shading: { fill: r[2] === "Hoog" ? "FDECEA" : r[2] === "Middel" ? "FFF3CD" : "EAF6EA", type: ShadingType.CLEAR },
            margins: { top: 80, bottom: 80, left: 120, right: 120 },
            children: [new Paragraph({ children: [new TextRun({ text: r[2], bold: true, size: 20, font: "Arial", color: r[2] === "Hoog" ? "C0392B" : r[2] === "Middel" ? "E67E22" : "27AE60" })] })]
          })
        ]
      }))
    ]
  });
}

// ── RAPPORT PARSER ─────────────────────────────────────────────────────────────
// Converteert de plain-text rapporttekst van Claude naar gestyled Word-content.
// Claude schrijft markdown-achtige structuur met ##, ###, -, bulletpunten, tabellen.
// Deze parser herkent die patronen en vertaalt ze naar de juiste docx-elementen.

function parseRapportTekst(tekst) {
  const regels = tekst.split('\n');
  const elementen = [];
  let i = 0;

  while (i < regels.length) {
    const regel = regels[i].trim();

    // Lege regel → kleine spatie
    if (!regel) {
      elementen.push(spacer(60));
      i++;
      continue;
    }

    // Markdown tabel detectie (regel die begint met |)
    if (regel.startsWith('|')) {
      const tabelRegels = [];
      while (i < regels.length && regels[i].trim().startsWith('|')) {
        const r = regels[i].trim();
        // Sla separator-rijen over (|---|---|)
        if (!r.match(/^\|[\s\-:]+\|/)) {
          tabelRegels.push(r);
        }
        i++;
      }
      if (tabelRegels.length > 0) {
        elementen.push(...parseTabel(tabelRegels));
      }
      continue;
    }

    // H1: # Tekst of 1. TEKST
    if (regel.match(/^#\s+/) || regel.match(/^\d+\.\s+[A-Z]/)) {
      const tekst = regel.replace(/^#+\s+/, '').replace(/^\d+\.\s+/, '');
      elementen.push(spacer(160));
      elementen.push(heading1(tekst));
      i++;
      continue;
    }

    // H2: ## Tekst of 5.1 Tekst
    if (regel.match(/^##\s+/) || regel.match(/^\d+\.\d+\s+/)) {
      const tekst = regel.replace(/^#+\s+/, '').replace(/^\d+\.\d+\s+/, '');
      elementen.push(spacer(80));
      elementen.push(heading2(tekst));
      i++;
      continue;
    }

    // H3: ### Tekst of #### Tekst
    if (regel.match(/^#{3,}\s+/)) {
      const tekst = regel.replace(/^#+\s+/, '');
      elementen.push(heading3(tekst));
      i++;
      continue;
    }

    // Bullet: - tekst of • tekst of * tekst
    if (regel.match(/^[-•*]\s+/)) {
      const tekst = regel.replace(/^[-•*]\s+/, '').replace(/\*\*(.*?)\*\*/g, '$1');
      elementen.push(bullet(tekst));
      i++;
      continue;
    }

    // Bold label + content: **Label**: tekst → infoBox
    if (regel.match(/^\*\*[^*]+\*\*\s*[:|]/)) {
      const match = regel.match(/^\*\*([^*]+)\*\*\s*[:|]\s*(.*)/);
      if (match) {
        elementen.push(infoBox(match[1].trim(), match[2].trim()));
        elementen.push(spacer(40));
        i++;
        continue;
      }
    }

    // Actiepunt tabel-blokken: herken actiepunt-secties op basis van omschrijving/aandachtsgebied/prioriteit
    if (regel.match(/^(Omschrijving|Aandachtsgebied|Prioriteit)\s*[:|]/i)) {
      // Collect action block
      const blok = {};
      while (i < regels.length && regels[i].trim()) {
        const r = regels[i].trim();
        const m = r.match(/^(Omschrijving|Aandachtsgebied|Prioriteit)\s*[:|]\s*(.*)/i);
        if (m) blok[m[1].toLowerCase()] = m[2];
        i++;
      }
      if (blok.omschrijving) {
        elementen.push(actionTable([[
          blok.omschrijving || '',
          blok.aandachtsgebied || '',
          blok.prioriteit || 'Middel'
        ]]));
        elementen.push(spacer(40));
      }
      continue;
    }

    // Gewone alinea tekst — strip markdown bold markers
    const schooneTekst = regel.replace(/\*\*(.*?)\*\*/g, '$1').replace(/\*(.*?)\*/g, '$1');
    elementen.push(body(schooneTekst));
    i++;
  }

  return elementen;
}

function parseTabel(regels) {
  if (regels.length === 0) return [];

  // Parse elke rij in cellen
  const rijen = regels.map(r =>
    r.split('|')
      .map(c => c.trim())
      .filter((c, idx, arr) => idx > 0 && idx < arr.length - 1)
  );

  if (rijen.length === 0) return [];

  const headers = rijen[0];
  const dataRijen = rijen.slice(1);
  const aantalKolommen = headers.length;

  // Bereken gelijke kolombreedtes
  const totaalBreedte = 9360;
  const kolBreedte = Math.floor(totaalBreedte / aantalKolommen);
  const kolBreedtes = Array(aantalKolommen).fill(kolBreedte);

  const maakHeaderCel = (tekst) => new TableCell({
    borders,
    shading: { fill: BLUE, type: ShadingType.CLEAR },
    margins: { top: 100, bottom: 100, left: 120, right: 120 },
    children: [new Paragraph({ children: [new TextRun({ text: tekst, bold: true, size: 20, font: "Arial", color: "FFFFFF" })] })]
  });

  const maakDataCel = (tekst, isGerij) => new TableCell({
    borders,
    shading: { fill: isGerij ? "F5F9FD" : "FFFFFF", type: ShadingType.CLEAR },
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    children: [new Paragraph({ children: [new TextRun({ text: tekst, size: 20, font: "Arial" })] })]
  });

  const tabelRijen = [
    new TableRow({ children: headers.map(h => maakHeaderCel(h)) }),
    ...dataRijen.map((rij, i) =>
      new TableRow({ children: rij.map(cel => maakDataCel(cel, i % 2 === 0)) })
    )
  ];

  return [
    new Table({
      width: { size: totaalBreedte, type: WidthType.DXA },
      columnWidths: kolBreedtes,
      rows: tabelRijen
    }),
    spacer(80)
  ];
}

// ── DOCUMENT BUILDER ───────────────────────────────────────────────────────────
function buildDocument(rapportTekst, onderzoekNaam, datum) {
  const inhoud = parseRapportTekst(rapportTekst);

  return new Document({
    numbering: {
      config: [
        {
          reference: "bullets",
          levels: [{
            level: 0,
            format: LevelFormat.BULLET,
            text: "•",
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 600, hanging: 300 } } }
          }]
        }
      ]
    },
    styles: {
      default: {
        document: { run: { font: "Arial", size: 22 } }
      },
      paragraphStyles: [
        {
          id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: 32, bold: true, font: "Arial", color: BLUE },
          paragraph: { spacing: { before: 400, after: 160 }, outlineLevel: 0 }
        },
        {
          id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: 26, bold: true, font: "Arial", color: MID_BLUE },
          paragraph: { spacing: { before: 280, after: 120 }, outlineLevel: 1 }
        },
        {
          id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: 22, bold: true, font: "Arial", color: "2C3E50" },
          paragraph: { spacing: { before: 200, after: 80 }, outlineLevel: 2 }
        }
      ]
    },
    sections: [{
      properties: {
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
        }
      },
      headers: {
        default: new Header({
          children: [
            new Paragraph({
              border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: MID_BLUE, space: 4 } },
              children: [
                new TextRun({ text: `${onderzoekNaam}`, size: 18, font: "Arial", color: GRAY }),
                new TextRun({ text: "    |    Vertrouwelijk", size: 18, font: "Arial", color: BORDER_COLOR })
              ]
            })
          ]
        })
      },
      footers: {
        default: new Footer({
          children: [
            new Paragraph({
              border: { top: { style: BorderStyle.SINGLE, size: 4, color: BORDER_COLOR, space: 4 } },
              children: [
                new TextRun({ text: `© Rinkel.com ${new Date().getFullYear()}  |  Intern gebruik  |  Pagina `, size: 18, font: "Arial", color: GRAY }),
                new TextRun({ children: [PageNumber.CURRENT], size: 18, font: "Arial", color: GRAY }),
                new TextRun({ text: " van ", size: 18, font: "Arial", color: GRAY }),
                new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 18, font: "Arial", color: GRAY }),
              ]
            })
          ]
        })
      },
      children: [
        // ── COVERPAGINA ──────────────────────────────────────────────────────
        new Paragraph({
          spacing: { before: 800, after: 200 },
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: "KLANTONDERZOEK — PRODUCT LED PERSPECTIEF", size: 20, font: "Arial", color: MID_BLUE, bold: true, allCaps: true })]
        }),
        new Paragraph({
          spacing: { before: 0, after: 160 },
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: "Wat moet het product doen om zichzelf te verkopen?", size: 44, font: "Arial", color: BLUE, bold: true })]
        }),
        new Paragraph({
          spacing: { before: 0, after: 40 },
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: onderzoekNaam, size: 28, font: "Arial", color: MID_BLUE, bold: true })]
        }),
        spacer(40),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { before: 0, after: 600 },
          children: [new TextRun({ text: `Gegenereerd door Claude AI  ·  ${datum}  ·  Vertrouwelijk`, size: 20, font: "Arial", color: GRAY })]
        }),

        // ── PLG INTRO BOX ────────────────────────────────────────────────────
        new Table({
          width: { size: 9360, type: WidthType.DXA },
          columnWidths: [9360],
          rows: [
            new TableRow({
              children: [
                new TableCell({
                  borders,
                  width: { size: 9360, type: WidthType.DXA },
                  shading: { fill: LIGHT_BLUE, type: ShadingType.CLEAR },
                  margins: { top: 180, bottom: 180, left: 200, right: 200 },
                  children: [
                    new Paragraph({ spacing: { before: 0, after: 80 }, children: [new TextRun({ text: "Over dit rapport", bold: true, size: 22, font: "Arial", color: BLUE })] }),
                    new Paragraph({ spacing: { before: 0, after: 0 }, children: [new TextRun({ text: "Dit rapport is geschreven vanuit een product led perspectief. Dat betekent dat acties en conclusies niet zijn geformuleerd vanuit sales of marketing, maar vanuit de vraag: wat moet het product doen om klanten vanzelf te laten converteren, blijven en groeien? Die vraag geldt voor elke afdeling — niet alleen voor het productteam.", size: 20, font: "Arial", color: "2C3E50" })] })
                  ]
                })
              ]
            })
          ]
        }),
        spacer(160),

        // ── RAPPORT INHOUD (gegenereerd door Claude) ─────────────────────────
        ...inhoud,

        // ── AFSLUITING ───────────────────────────────────────────────────────
        spacer(120),
        body("— Einde rapport —", { color: GRAY, italics: true }),
        spacer(40),
      ]
    }]
  });
}

// ── EXPRESS ENDPOINT ───────────────────────────────────────────────────────────
app.post('/generate', async (req, res) => {
  try {
    const { rapportTekst, onderzoekNaam, datum } = req.body;

    if (!rapportTekst) {
      return res.status(400).json({ error: 'rapportTekst is verplicht' });
    }

    const naamLabel = onderzoekNaam || 'Rinkel Onderzoeksrapport';
    const datumLabel = datum || new Date().toLocaleDateString('nl-NL', { day: 'numeric', month: 'long', year: 'numeric' });

    const doc = buildDocument(rapportTekst, naamLabel, datumLabel);
    const buffer = await Packer.toBuffer(doc);

    res.set({
      'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      'Content-Disposition': `attachment; filename="${naamLabel.replace(/[^a-zA-Z0-9]/g, '_')}.docx"`,
      'Content-Length': buffer.length
    });

    res.send(buffer);
    console.log(`[OK] Rapport gegenereerd: ${naamLabel} (${buffer.length} bytes)`);

  } catch (err) {
    console.error('[FOUT]', err);
    res.status(500).json({ error: 'Fout bij genereren rapport', detail: err.message });
  }
});

// Health check voor Render
app.get('/', (req, res) => res.json({ status: 'ok', service: 'Rinkel Rapport Generator' }));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Rapport server draait op poort ${PORT}`));
