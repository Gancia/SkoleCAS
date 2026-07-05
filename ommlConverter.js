/**
 * ommlConverter.js  v38
 * Pipeline: LaTeX → MathML (KaTeX) → OMML (XSLT) → Flat OPC OOXML (Word)
 *
 * Rettelser i v28:
 *  1. KaTeX mathml-output re-parses via DOMParser("text/xml") så MathML-namespace
 *     er korrekt sat — setAttribute("xmlns",...) virker IKKE på HTML-parsede noder.
 *  2. JS-fallback OMML-konverter dækker \\frac, ^{}, \\sqrt, \\sin/cos/tan osv.,
 *     og hele tal / decimaler — bruges hvis XSLT ikke loader.
 *  3. XSLT-fejl logger nu tydeligt; initReady sættes korrekt.
 *  4. latexToOoxml() venter (via retry) til KaTeX er loaded (defer-race).
 */

// ─────────────────────────────────────────────────────────────────
//  State
// ─────────────────────────────────────────────────────────────────
let xsltProcessor = null;
window.ommlConverterReady = false;
window.ommlConverterError = null;

// ─────────────────────────────────────────────────────────────────
//  XSLT loader
// ─────────────────────────────────────────────────────────────────
async function initOmmlConverter() {
    try {
        const response = await fetch("mml2omml.xsl");
        if (!response.ok) {
            throw new Error(`HTTP ${response.status} ved hentning af mml2omml.xsl`);
        }
        const xsltText = await response.text();
        const parser = new DOMParser();
        const xslStylesheet = parser.parseFromString(xsltText, "application/xml");

        const parseErr = xslStylesheet.querySelector("parsererror");
        if (parseErr) {
            throw new Error("mml2omml.xsl parse-fejl: " + parseErr.textContent.substring(0, 200));
        }

        xsltProcessor = new XSLTProcessor();
        xsltProcessor.importStylesheet(xslStylesheet);
        window.ommlConverterReady = true;
        console.log("✅ OMML Converter (XSLT) klar!");
    } catch (e) {
        console.error("❌ XSLT-loader fejlede — JS-fallback aktiveres:", e);
        window.ommlConverterError = e.message;
        // ommlConverterReady forbliver false — latexToOoxml bruger JS-fallback
    }
}

initOmmlConverter();

// ─────────────────────────────────────────────────────────────────
//  sanitizeLatex
// ─────────────────────────────────────────────────────────────────
function sanitizeLatex(latex) {
    return latex
        .replace(/\\dfrac/g, '\\frac')
        .replace(/\\operatorname\{(sin|cos|tan|log|ln|exp|abs|sqrt|cot|sec|csc)\}/g, '\\$1')
        .replace(/\\left\(([^()\\]+)\\right\)/g, '($1)')
        .replace(/  +/g, ' ')
        .trim();
}

// ─────────────────────────────────────────────────────────────────
//  JS-FALLBACK
//  (Fjernet JS-fallback ifølge plan for at sikre streng OOXML-compliance)
// ─────────────────────────────────────────────────────────────────

// ─────────────────────────────────────────────────────────────────
//  XSLT-pipeline: LaTeX → MathML (KaTeX) → OMML (XSLT)
// ─────────────────────────────────────────────────────────────────
function latexToOmmlViaXslt(latexStr) {
    const sanitized = sanitizeLatex(latexStr);

    // 1. LaTeX → MathML string via KaTeX
    let mathmlString;
    try {
        mathmlString = window.katex.renderToString(sanitized, {
            output: 'mathml',
            throwOnError: true,
            displayMode: false
        });
    } catch (katexErr) {
        console.error('❌ KaTeX parse-fejl. Input:', JSON.stringify(sanitized));
        throw new Error(`KaTeX: ${katexErr.message} (LaTeX: ${sanitized.substring(0, 60)})`);
    }

    // 2. Udtræk <math>...</math> streng (KaTeX returnerer <span><math>...</math></span>)
    //    Vi SKAL re-parse med "text/xml" for at MathML-namespace er korrekt sat.
    //    DOMParser("text/html") sætter IKKE MathML-namespace, og
    //    setAttribute("xmlns",...) virker ikke på HTML-parsede noder.
    const mathMatch = mathmlString.match(/<math[\s\S]*<\/math>/);
    if (!mathMatch) {
        throw new Error("KaTeX returnerede ingen <math>-element.");
    }

    // Tilføj eksplicit namespace-deklaration i strengen så XML-parseren forstår den
    let mathXml = mathMatch[0];
    if (!mathXml.includes('xmlns=')) {
        mathXml = mathXml.replace('<math', '<math xmlns="http://www.w3.org/1998/Math/MathML"');
    }

    const xmlParser = new DOMParser();
    const mathDoc = xmlParser.parseFromString(mathXml, "text/xml");

    const parseErr = mathDoc.querySelector("parsererror");
    if (parseErr) {
        console.error("❌ MathML XML-parse fejlede:", parseErr.textContent);
        throw new Error("MathML kunne ikke parses som XML: " + parseErr.textContent.substring(0, 120));
    }

    const mathNode = mathDoc.documentElement;

    // FJERNER ANNOTATIONS!
    // KaTeX indsætter <annotation> med rå LaTeX (fx "\frac{1}{2}"). 
    const annotations = mathDoc.querySelectorAll("annotation");
    for (let i = 0; i < annotations.length; i++) {
        annotations[i].parentNode.removeChild(annotations[i]);
    }
    
    // Udpak <semantics> tagget for at sikre ren MathML før XSLT processerer det,
    // og derved forhindre uønsket HTML/XML leakage.
    const semantics = mathDoc.querySelectorAll("semantics");
    for (let i = 0; i < semantics.length; i++) {
        const sem = semantics[i];
        while (sem.firstChild) {
            sem.parentNode.insertBefore(sem.firstChild, sem);
        }
        sem.parentNode.removeChild(sem);
    }

    // 3. MathML → OMML via XSLT
    const ommlDoc = xsltProcessor.transformToDocument(mathNode);
    const serializer = new XMLSerializer();
    let ommlString = serializer.serializeToString(ommlDoc);

    // 4. Rens serialiseret OMML (Fjern kun selve XML deklarationen, behold namespaces)
    ommlString = ommlString.replace(/<\?xml[^?]*\?>\s*/g, '').trim();

    return ommlString;
}

// ─────────────────────────────────────────────────────────────────
//  buildFlatOpc — pakker et OMML-fragment i en Flat OPC Word-pakke
// ─────────────────────────────────────────────────────────────────
function buildFlatOpc(ommlString) {
    return `<?xml version="1.0" standalone="yes"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/word/document.xml"
            pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document
        xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <w:body>
          <w:p>
            ${ommlString}
          </w:p>
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`;
}

// ─────────────────────────────────────────────────────────────────
//  latexToOoxml — offentlig API
// ─────────────────────────────────────────────────────────────────
window.latexToOoxml = function(latexStr) {
    const katexReady = (typeof window.katex !== 'undefined');

    if (!katexReady) {
        throw new Error("KaTeX ikke loaded endnu.");
    }
    if (!xsltProcessor) {
        throw new Error("XSLT processor ikke loaded: " + window.ommlConverterError);
    }

    const omml = latexToOmmlViaXslt(latexStr);
    return buildFlatOpc(omml);
};

window.latexToOmmlString = function(latexStr) {
    const katexReady = (typeof window.katex !== 'undefined');
    
    if (!katexReady) {
        throw new Error("KaTeX ikke loaded endnu.");
    }
    if (!xsltProcessor) {
        throw new Error("XSLT processor ikke loaded: " + window.ommlConverterError);
    }

    return latexToOmmlViaXslt(latexStr);
};

// ─────────────────────────────────────────────────────────────────
//  formatFractions — konverterer mellem / og \frac baseret på indstilling
// ─────────────────────────────────────────────────────────────────
window.formatFractions = function(fullLatex, isStacked) {
    if (isStacked) {
        // Erstat 1/2 og (a+b)/c med \frac{1}{2} og \frac{a+b}{c}
        let result = fullLatex;
        let index = result.indexOf('/');
        while (index !== -1) {
            let leftStart = index - 1;
            let num = '';
            if (result[leftStart] === ')') {
                let depth = 1; leftStart--;
                while (leftStart >= 0 && depth > 0) {
                    if (result[leftStart] === ')') depth++;
                    if (result[leftStart] === '(') depth--;
                    leftStart--;
                }
                leftStart++;
                num = result.substring(leftStart + 1, index - 1);
            } else {
                while (leftStart >= 0 && /[a-zA-Z0-9_.\\{}]/.test(result[leftStart])) leftStart--;
                leftStart++;
                num = result.substring(leftStart, index);
            }
            
            let rightEnd = index + 1;
            let den = '';
            if (result[rightEnd] === '(') {
                let depth = 1; rightEnd++;
                while (rightEnd < result.length && depth > 0) {
                    if (result[rightEnd] === '(') depth++;
                    if (result[rightEnd] === ')') depth--;
                    rightEnd++;
                }
                rightEnd--;
                den = result.substring(index + 2, rightEnd);
            } else {
                while (rightEnd < result.length && /[a-zA-Z0-9_.\\{}]/.test(result[rightEnd])) rightEnd++;
                rightEnd--;
                den = result.substring(index + 1, rightEnd + 1);
            }
            
            if (leftStart < index && rightEnd > index && num && den) {
                result = result.substring(0, leftStart) + `\\frac{${num}}{${den}}` + result.substring(rightEnd + 1);
                index = result.indexOf('/');
            } else {
                index = result.indexOf('/', index + 1);
            }
        }
        return result;
    } else {
        // Konverter \frac{a}{b} til a/b
        let result = fullLatex;
        let index = result.indexOf('\\frac{');
        while (index !== -1) {
            let p1Start = index + 5;
            if (result[p1Start] !== '{') break;
            let depth1 = 1, p1End = p1Start + 1;
            while (p1End < result.length && depth1 > 0) {
                if (result[p1End] === '{') depth1++;
                if (result[p1End] === '}') depth1--;
                p1End++;
            }
            p1End--;
            let num = result.substring(p1Start + 1, p1End);
            
            let p2Start = p1End + 1;
            if (result[p2Start] !== '{') break;
            let depth2 = 1, p2End = p2Start + 1;
            while (p2End < result.length && depth2 > 0) {
                if (result[p2End] === '{') depth2++;
                if (result[p2End] === '}') depth2--;
                p2End++;
            }
            p2End--;
            let den = result.substring(p2Start + 1, p2End);
            
            if (/[\+\-\*\=\s]/.test(num) && !(num.startsWith('(') && num.endsWith(')'))) num = `(${num})`;
            if (/[\+\-\*\=\s]/.test(den) && !(den.startsWith('(') && den.endsWith(')'))) den = `(${den})`;
            
            result = result.substring(0, index) + `${num}/${den}` + result.substring(p2End + 1);
            index = result.indexOf('\\frac{');
        }
        return result;
    }
};
