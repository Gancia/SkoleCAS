const fs = require('fs');

// Stub window and DOMParser
global.window = {};

const ommlConverterContent = fs.readFileSync('ommlConverter.js', 'utf8');

// faking tokenizeLatex, etc. Actually, I can just evaluate the script.
eval(ommlConverterContent.replace('initOmmlConverter();', '')); // skip async fetch

const testExpr = '\\frac{1}{2}';
const out = latexToOoxml(testExpr);

console.log(out);
