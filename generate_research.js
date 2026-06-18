const { Document, Packer, Paragraph, TextRun, HeadingLevel } = require("docx");
const fs = require("fs");

const doc = new Document({
    sections: [{
        children: [
            new Paragraph({
                text: "Philippine Culture: A Rich Tapestry",
                heading: HeadingLevel.HEADING_1,
            }),
            new Paragraph({
                text: "Introduction",
                heading: HeadingLevel.HEADING_2,
            }),
            new Paragraph({
                text: "The culture of the Philippines is a vibrant blend of indigenous traditions, Spanish colonial heritage, and American influence. This unique combination has shaped the Filipino identity, characterized by resilience, hospitality, and strong family ties.",
            }),
            new Paragraph({
                text: "Key Aspects of Philippine Culture",
                heading: HeadingLevel.HEADING_2,
            }),
            new Paragraph({
                text: "Language: The Philippines is home to over 170 languages, with Filipino (based on Tagalog) and English serving as official languages.",
                bullet: { level: 0 }
            }),
            new Paragraph({
                text: "Festivals: Filipinos are known for their colorful festivals or 'fiestas', which celebrate patron saints, bountiful harvests, and historical events, such as the Sinulog and Ati-Atihan.",
                bullet: { level: 0 }
            }),
            new Paragraph({
                text: "Cuisine: Filipino food is a flavorful mix of sweet, sour, and salty, with iconic dishes like Adobo, Sinigang, and the popular dessert, Halo-halo.",
                bullet: { level: 0 }
            }),
            new Paragraph({
                text: "Values: A core value in Philippine culture is 'Bayanihan', which refers to communal unity and helping one another without expecting anything in return.",
                bullet: { level: 0 }
            }),
            new Paragraph({
                text: "Conclusion",
                heading: HeadingLevel.HEADING_2,
            }),
            new Paragraph({
                text: "Philippine culture continues to evolve, reflecting the dynamic spirit of its people. It remains a testament to the nation's rich history and diverse influences, making it truly unique in Southeast Asia.",
            }),
        ],
    }],
});

Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync("output/PhilippineCultureResearch.docx", buffer);
    console.log("Research document created successfully!");
});
