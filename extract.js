import fs from 'fs';
import pdf from 'pdf-parse';

const dataBuffer = fs.readFileSync('./Prueba1.pdf');

pdf(dataBuffer).then(function(data) {
    fs.writeFileSync('./Prueba1_text.txt', data.text);
    console.log("PDF parsed successfully! Text length:", data.text.length);
}).catch(err => {
    console.error("Error parsing PDF:", err);
});
