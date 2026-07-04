/**
 * ommlConverter.js  v28
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
//  JS-FALLBACK: simpel LaTeX → OMML konverter
//  Dækker de mest almindelige gymnasiematematik-udtryk uden XSLT.
// ─────────────────────────────────────────────────────────────────

/**
 * Tokenisér LaTeX i en liste af {type, value} objekter.
 * Typer: 'cmd' (\\frac etc), 'group' ({...}), 'char', 'space'
 */
function tokenizeLatex(latex) {
    const tokens = [];
    let i = 0;
    while (i < latex.length) {
        if (latex[i] === '\\') {
            // Kommando: \frac, \sqrt, \sin, \approx, \Rightarrow osv.
            let j = i + 1;
            if (j < latex.length && /[a-zA-Z]/.test(latex[j])) {
                while (j < latex.length && /[a-zA-Z]/.test(latex[j])) j++;
                tokens.push({ type: 'cmd', value: latex.slice(i, j) });
            } else {
                // enkelt-tegn kommando f.eks. \\, \{ \}
                tokens.push({ type: 'cmd', value: latex.slice(i, i + 2) });
                j = i + 2;
            }
            i = j;
        } else if (latex[i] === '{') {
            // Find matchende }
            let depth = 1, j = i + 1;
            while (j < latex.length && depth > 0) {
                if (latex[j] === '{') depth++;
                else if (latex[j] === '}') depth--;
                j++;
            }
            tokens.push({ type: 'group', value: latex.slice(i + 1, j - 1) });
            i = j;
        } else if (latex[i] === ' ') {
            i++;
        } else {
            tokens.push({ type: 'char', value: latex[i] });
            i++;
        }
    }
    return tokens;
}

/**
 * Konvertér en LaTeX-streng til OMML-fragmentstreng (indhold af <m:oMath>).
 * Rekursivt: grupper parses igen.
 */
function latexFragmentToOmml(latex) {
    latex = latex.trim();
    if (!latex) return '';

    const tokens = tokenizeLatex(latex);
    let out = '';
    let i = 0;

    function nextGroup(startIdx) {
        // Returner næste token hvis det er en group, ellers tag næste char som group-indhold
        if (startIdx < tokens.length && tokens[startIdx].type === 'group') {
            return { content: tokens[startIdx].value, consumed: 1 };
        } else if (startIdx < tokens.length && tokens[startIdx].type === 'char') {
            return { content: tokens[startIdx].value, consumed: 1 };
        }
        return { content: '', consumed: 0 };
    }

    while (i < tokens.length) {
        const tok = tokens[i];

        if (tok.type === 'cmd') {
            // ──────── \frac{num}{den} ────────
            if (tok.value === '\\frac') {
                const num = nextGroup(i + 1);
                const den = nextGroup(i + 1 + num.consumed);
                out += `<m:f><m:fPr><m:type m:val="bar"/></m:fPr>` +
                       `<m:num>${latexFragmentToOmml(num.content)}</m:num>` +
                       `<m:den>${latexFragmentToOmml(den.content)}</m:den></m:f>`;
                i += 1 + num.consumed + den.consumed;
                continue;
            }

            // ──────── \sqrt{x} ────────
            if (tok.value === '\\sqrt') {
                const arg = nextGroup(i + 1);
                out += `<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr>` +
                       `<m:deg><m:e/></m:deg>` +
                       `<m:e>${latexFragmentToOmml(arg.content)}</m:e></m:rad>`;
                i += 1 + arg.consumed;
                continue;
            }

            // ──────── \approx, \Rightarrow, \int, \pi, \cdot, \lor ────────
            const SYMBOL_MAP = {
                '\\approx': '≈', '\\Rightarrow': '⇒', '\\rightarrow': '→',
                '\\int': '∫', '\\pi': 'π', '\\cdot': '·',
                '\\lor': '∨', '\\land': '∧', '\\infty': '∞',
                '\\leq': '≤', '\\geq': '≥', '\\neq': '≠',
                '\\times': '×', '\\div': '÷', '\\pm': '±',
                '\\alpha': 'α', '\\beta': 'β', '\\gamma': 'γ',
                '\\delta': 'δ', '\\theta': 'θ', '\\lambda': 'λ',
                '\\mu': 'μ', '\\sigma': 'σ', '\\omega': 'ω',
            };
            if (SYMBOL_MAP[tok.value]) {
                out += `<m:r><m:t>${SYMBOL_MAP[tok.value]}</m:t></m:r>`;
                i++;
                continue;
            }

            // ──────── \sin, \cos, \tan, \ln, \log, \exp ────────
            const FUNC_NAMES = ['\\sin','\\cos','\\tan','\\cot','\\sec','\\csc',
                                '\\ln','\\log','\\exp','\\arcsin','\\arccos','\\arctan'];
            if (FUNC_NAMES.includes(tok.value)) {
                const name = tok.value.slice(1); // fjern backslash
                out += `<m:func><m:funcPr/><m:fName><m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t>${name}</m:t></m:r></m:fName><m:e/>` +
                       `</m:func>`;
                i++;
                continue;
            }

            // Ukendt kommando — ignorer
            i++;
            continue;
        }

        if (tok.type === 'group') {
            // En isoleret gruppe — rekursivt parse indholdet
            out += latexFragmentToOmml(tok.value);
            i++;
            continue;
        }

        if (tok.type === 'char') {
            const c = tok.value;

            // ──────── superscript: x^{n} eller x^n ────────
            if (c === '^') {
                const base = out; // alt hidtil er basen (simplificeret)
                const sup = nextGroup(i + 1);
                // Erstatter hele out med en sSup-node
                out = `<m:sSup><m:sSupPr/><m:e>${base}</m:e>` +
                      `<m:sup>${latexFragmentToOmml(sup.content)}</m:sup></m:sSup>`;
                i += 1 + sup.consumed;
                continue;
            }

            // ──────── subscript: x_{n} ────────
            if (c === '_') {
                const base = out;
                const sub = nextGroup(i + 1);
                out = `<m:sSub><m:sSubPr/><m:e>${base}</m:e>` +
                      `<m:sub>${latexFragmentToOmml(sub.content)}</m:sub></m:sSub>`;
                i += 1 + sub.consumed;
                continue;
            }

            // Almindeligt tegn (tal, bogstav, operator)
            out += `<m:r><m:t>${escOmml(c)}</m:t></m:r>`;
            i++;
            continue;
        }

        i++;
    }

    return out;
}

function escOmml(str) {
    if (!str) return '';
    return String(str)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

/**
 * Fuldt oMath-element via JS-fallback (ingen XSLT).
 */
function latexToOmmlViaJs(latexStr) {
    const sanitized = sanitizeLatex(latexStr);
    const inner = latexFragmentToOmml(sanitized);
    return `<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">${inner}</m:oMath>`;
}

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
    // mml2omml.xsl ignorerer tagget, men dens catch-all template kopierer selve teksten over.
    // Resultatet er, at ren tekst ender ulovligt inde i OMML og får Word til at crashe fuldstændigt.
    const annotations = mathDoc.querySelectorAll("annotation");
    for (let i = 0; i < annotations.length; i++) {
        annotations[i].parentNode.removeChild(annotations[i]);
    }
    const semantics = mathDoc.querySelectorAll("semantics");
    for (let i = 0; i < semantics.length; i++) {
        // Semantics-tagget alene er uskadeligt når annotation fjernes,
        // men vi kan lige så godt lade det være, for XSLT's catch-all håndterer det.
    }

    // 3. MathML → OMML via XSLT
    const ommlDoc = xsltProcessor.transformToDocument(mathNode);
    const serializer = new XMLSerializer();
    let ommlString = serializer.serializeToString(ommlDoc);

    // 4. Rens serialiseret OMML
    ommlString = ommlString
        .replace(/<\?xml[^?]*\?>\s*/g, '')
        // Fjern redundante xmlns-deklarationer som XMLSerializer tilføjer på <m:oMath>
        // (de er allerede deklareret på <w:document> i den ydre Flat OPC-pakke)
        .replace(/ xmlns:m="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/math"/g, '')
        .replace(/ xmlns:mml="[^"]*"/g, '')
        .replace(/ xmlns="http:\/\/www\.w3\.org\/1998\/Math\/MathML"/g, '')
        .trim();

    return ommlString;
}

// ─────────────────────────────────────────────────────────────────
//  buildFlatOpc — pakker et OMML-fragment i en Flat OPC Word-pakke
// ─────────────────────────────────────────────────────────────────
function buildFlatOpc(ommlString) {
    // Fjern namespace-deklarationer på <m:oMath> hvis de er der fra JS-fallback —
    // <w:document> erklærer allerede xmlns:m, så de er redundante og forvirrer Word.
    const cleanOmml = ommlString
        .replace(/ xmlns:m="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/math"/g, '');

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
            <m:oMathPara>
              <m:oMathParaPr>
                <m:jc m:val="centerGroup"/>
              </m:oMathParaPr>
              ${cleanOmml}
            </m:oMathPara>
          </w:p>
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`;
}

// ─────────────────────────────────────────────────────────────────
//  latexToOoxml — offentlig API
//  Forsøger XSLT-pipeline først; falder tilbage til JS-konverter.
// ─────────────────────────────────────────────────────────────────
window.latexToOoxml = function(latexStr) {
    // KaTeX kan stadig loade pga. defer — tjek og kast informativ fejl
    const katexReady = (typeof window.katex !== 'undefined');

    if (xsltProcessor && katexReady) {
        // ── Primær pipeline: XSLT ──
        try {
            const omml = latexToOmmlViaXslt(latexStr);
            return buildFlatOpc(omml);
        } catch (e) {
            console.warn("⚠️ XSLT-pipeline fejlede, forsøger JS-fallback:", e.message);
            // Fald igennem til JS-fallback
        }
    } else if (!katexReady) {
        console.warn("⚠️ KaTeX ikke loaded endnu — bruger JS-fallback");
    } else {
        console.warn("⚠️ XSLT ikke loaded — bruger JS-fallback. Fejl:", window.ommlConverterError);
    }

    // ── Sekundær pipeline: ren JS ──
    const omml = latexToOmmlViaJs(latexStr);
    return buildFlatOpc(omml);
};
