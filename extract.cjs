const fs = require('fs');
const pdfImport = require('pdf-parse');

async function run() {
    const dataBuffer = fs.readFileSync('./Prueba1.pdf');
    const uint8Array = new Uint8Array(dataBuffer);
    const parser = new pdfImport.PDFParse(uint8Array);
    
    const resolved = await parser.getText();
    const text = resolved.text;
    
    fs.writeFileSync('./Prueba1_text.txt', text);
    console.log("PDF parsed successfully! Text length:", text.length);
}

run().catch(console.error);
