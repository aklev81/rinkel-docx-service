const express = require('express');
const { Document, Packer, Paragraph, TextRun, BorderStyle, WidthType, ShadingType } = require('docx');

const app = express();
app.use(express.json({ limit: '10mb' }));

const BLAUW = "1B3A6B";
const ROOD = "E8222A";

function h1(text) {
  return new Paragraph({
    spacing: { before: 280, after: 120 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: BLAUW, space: 3 } },
    children: [new TextRun({ text, bold: true, size: 26, color: BLAUW, font: "Arial" })]
  });
}

function h2(text) {
  return new Paragraph({
    spacing: { before: 180, after: 80 },
    children: [new TextRun({ text, bold: true, size: 22, color: "1a1a1a", font: "Arial" })]
  });
}

function body(text) {
  return new Paragraph({
    spacing: { before: 0, after: 80 },
    children: [new TextRun({ text: text.replace(/\*\*/g, ''), size: 20, font: "Arial", color: "1a1a1a" })]
  });
}

function bullet(text) {
  return new Paragraph({
    indent: { left: 360, hanging: 180 },
    spacing: { before: 0, after: 60 },
    children: [new TextRun({ text: `• ${text.replace(/^[-*] /, '')}`, size: 20, font: "Arial", color: "1a1a1a" })]
  });
}

app.post('/generate', async (req, res) => {
  try {
    const { rapportTekst, onderzoekNaam, datum } = req.body;
    const content = [];

    content.push(new Paragraph({
      spacing: { before: 0, after: 60 },
      children: [new TextRun({ text: "rinkel", bold: true, size: 44, color: ROOD, font: "Arial" })]
    }));
    content.push(new Paragraph({
      border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: BLAUW, space: 4 } },
      spacing: { before: 0, after: 0 },
      children: []
    }));
    content.push(new Paragraph({
      spacing: { before: 120, after: 60 },
      children: [new TextRun({ text: onderzoekNaam || 'Onderzoeksrapport', bold: true, size: 32, color: "1a1a1a", font: "Arial" })]
    }));
    content.push(new Paragraph({
      spacing: { before: 0, after: 240 },
      children: [new TextRun({ text: `Gegenereerd door Claude AI  ·  ${datum || new Date().toLocaleDateString('nl-NL')}  ·  Vertrouwelijk`, size: 18, color: "666666", font: "Arial" })]
    }));

    const regels = rapportTekst.split('\n').filter(r => r.trim());
    for (const regel of regels) {
      if (regel.startsWith('# ')) content.push(h1(regel.replace(/^# /, '')));
      else if (regel.startsWith('## ')) content.push(h2(regel.replace(/^## /, '')));
      else if (regel.startsWith('### ')) content.push(new Paragraph({
        spacing: { before: 120, after: 60 },
        children: [new TextRun({ text: regel.replace(/^### /, ''), bold: true, size: 20, color: BLAUW, font: "Arial" })]
      }));
      else if (regel.startsWith('- ') || regel.startsWith('* ')) content.push(bullet(regel));
      else if (regel.trim()) content.push(body(regel));
    }

    const doc = new Document({
      sections: [{
        properties: {
          page: {
            size: { width: 11906, height: 16838 },
            margin: { top: 1134, right: 1134, bottom: 1134, left: 1134 }
          }
        },
        children: content
      }]
    });

    const buffer = await Packer.toBuffer(doc);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="${onderzoekNaam || 'rapport'}.docx"`);
    res.send(buffer);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.get('/', (req, res) => res.send('Rinkel DOCX Service actief'));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Service draait op poort ${PORT}`));
