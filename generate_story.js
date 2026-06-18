const { Document, Packer, Paragraph, TextRun, HeadingLevel } = require("docx");
const fs = require("fs");

const doc = new Document({
    sections: [{
        properties: {},
        children: [
            new Paragraph({
                text: "Ang Batang Mahilig Magbasa",
                heading: HeadingLevel.HEADING_1,
            }),
            new Paragraph({
                children: [
                    new TextRun("Ang maliit na si Maria ay mahilig magbasa. Sa ilalim ng puno ng mangga, hawak niya ang kanyang paboritong aklat. \"Balang araw, lilibutin ko ang buong mundo,\" bulong niya sa kanyang sarili. Kahit simple lang ang kanilang buhay, pangarap niya ay abot-langit."),
                ],
                spacing: { after: 200 },
            }),
        ],
    }],
});

Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync("output/AngBatangMahiligMagbasa.docx", buffer);
    console.log("Document created successfully!");
});
