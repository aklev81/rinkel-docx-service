const express = require('express');
const { Document, Packer, Paragraph, TextRun, BorderStyle } = require('docx');
const fs = require('fs');
const path = require('path');

const app = express();
app.use(express.json({ limit: '10mb' }));

const BLAUW = "1B3A6B";
const ROOD = "E8222A";
const UPLOAD_DIR = path.join(__dirname, 'uploads');

if (!fs.existsSync(UPLOAD_DIR)) fs.mkdirSync(UPLOAD_DIR);

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
    const fileName = `${(onderzoekNaam || 'rapport').replace(/[^a-zA-Z0-9]/g, '_')}_${Date.now()}.docx`;
    const filePath = path.join(UPLOAD_DIR, fileName);
    fs.writeFileSync(filePath, buffer);

    const baseUrl = process.env.BASE_URL || `https://express-hello-world-8h0x.onrender.com`;
    const downloadUrl = `${baseUrl}/download/${fileName}`;

    res.json({ success: true, downloadUrl, fileName });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.get('/download/:fileName', (req, res) => {
  const filePath = path.join(UPLOAD_DIR, req.params.fileName);
  if (!fs.existsSync(filePath)) return res.status(404).json({ error: 'File not found' });
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
  res.setHeader('Content-Disposition', `attachment; filename="${req.params.fileName}"`);
  res.send(fs.readFileSync(filePath));
});

app.get('/', (req, res) => res.send('Rinkel DOCX Service actief'));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Service draait op poort ${PORT}`));
