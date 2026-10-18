/**
 * ommlConverter.js — To-vejs Native JS OMML-Bro (0 XSLT / 0 mml2omml.xsl)
 *
 * Indeholder to kerne-oversættere:
 *  1. IND I WORD (LaTeX -> OMML AST-kompiler i ren JS):
 *     Bygger direkte <m:oMath> XML fra LaTeX uden XSLTProcessor eller mml2omml.xsl.
 *     Understøtter <m:f> (brøker), <m:rad> (rødder), <m:sSup>/<m:sSub> (potenser/indices),
 *     <m:d> (skalerende parenteser), <m:func> (funktioner) og danske decimaltal (3,14).
 *
 *  2. UD AF WORD (OMML XML -> Matematisk udtryk):
 *     Parser rigtige Word-ligninger (<m:oMath> fra Words ligningseditor Alt+= eller
 *     tidligere SkoleCAS-resultater) tilbage til et regnbart udtryk, præcis som WordMat.
 */

window.ommlConverterReady = true;
window.ommlConverterError = null;

window.ensureOmmlReady = async function() {
    return true;
};

// ─────────────────────────────────────────────────────────────────
//  Hjælpefunktioner til XML
// ─────────────────────────────────────────────────────────────────

function escapeXml(str) {
    return String(str)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

function makeRun(text, isNormalText) {
    if (!text) return "";
    const rPr = isNormalText ? '<m:rPr><m:sty m:val="p"/></m:rPr>' : '';
    return `<m:r>${rPr}<m:t>${escapeXml(text)}</m:t></m:r>`;
}

/**
 * Tjekker om en streng er fuldt omsluttet af ét matchende ydre parentespar.
 */
function isFullyWrappedInParens(str) {
    if (!str || str.length < 2 || str[0] !== '(' || str[str.length - 1] !== ')') {
        return false;
    }
    let depth = 0;
    for (let i = 0; i < str.length; i++) {
        const ch = str[i];
        if (ch === '(') depth++;
        else if (ch === ')') depth--;
        if (depth === 0 && i < str.length - 1) {
            return false;
        }
    }
    return depth === 0;
}

// ─────────────────────────────────────────────────────────────────
//  DEL 1: LaTeX -> OMML Kompiler (Ren JS, 0 XSLT)
// ─────────────────────────────────────────────────────────────────

const LATEX_SYMBOLS = {
    '\\pi': 'π',
    '\\Pi': 'Π',
    '\\alpha': 'α',
    '\\beta': 'β',
    '\\gamma': 'γ',
    '\\delta': 'δ',
    '\\theta': 'θ',
    '\\lambda': 'λ',
    '\\sigma': 'σ',
    '\\omega': 'ω',
    '\\infty': '∞',
    '\\approx': ' ≈ ',
    '\\Rightarrow': ' ⇒ ',
    '\\rightarrow': ' → ',
    '\\leftrightarrow': ' ⇔ ',
    '\\Leftrightarrow': ' ⇔ ',
    '\\lor': ' ∨ ',
    '\\land': ' ∧ ',
    '\\pm': ' ± ',
    '\\mp': ' ∓ ',
    '\\cdot': '⋅',
    '\\times': '×',
    '\\div': '÷',
    '\\leq': ' ≤ ',
    '\\le': ' ≤ ',
    '\\geq': ' ≥ ',
    '\\ge': ' ≥ ',
    '\\neq': ' ≠ ',
    '\\ne': ' ≠ ',
    '\\int': '∫',
    '\\quad': '  ',
    '\\qquad': '    ',
    '\\,': ' ',
    '\\;': ' ',
    '\\:': ' '
};

const UPRIGHT_FUNCTIONS = new Set([
    'sin', 'cos', 'tan', 'cot', 'sec', 'csc',
    'arcsin', 'arccos', 'arctan', 'asin', 'acos', 'atan',
    'sinh', 'cosh', 'tanh', 'ln', 'log', 'exp', 'abs', 'lim', 'max', 'min'
]);

function sanitizeLatex(latex) {
    if (!latex || typeof latex !== 'string') return "";
    return latex
        .replace(/\\dfrac/g, '\\frac')
        .replace(/\\tfrac/g, '\\frac')
        .replace(/\\operatorname\{([^}]+)\}/g, '\\$1')
        .replace(/\\left\s*([()[\]|])/g, '$1')
        .replace(/\\right\s*([()[\]|])/g, '$1')
        .replace(/  +/g, ' ')
        .trim();
}

/**
 * Læser en LaTeX-gruppe { ... } eller et enkelt token fra pos i str.
 * Returnerer { content: string, nextPos: number }.
 */
function readLatexGroupOrToken(str, pos) {
    while (pos < str.length && str[pos] === ' ') pos++;
    if (pos >= str.length) return { content: "", nextPos: pos };

    if (str[pos] === '{') {
        let depth = 1;
        let i = pos + 1;
        while (i < str.length && depth > 0) {
            if (str[i] === '{' && str[i - 1] !== '\\') depth++;
            else if (str[i] === '}' && str[i - 1] !== '\\') depth--;
            i++;
        }
        return {
            content: str.substring(pos + 1, i - 1),
            nextPos: i
        };
    }

    if (str[pos] === '\\') {
        let i = pos + 1;
        while (i < str.length && /[a-zA-Z]/.test(str[i])) i++;
        if (i === pos + 1 && i < str.length) i++;
        return {
            content: str.substring(pos, i),
            nextPos: i
        };
    }

    return {
        content: str[pos],
        nextPos: pos + 1
    };
}

/**
 * Finder matchende lukke-parentes ')' eller ']' fra startPos (som peger på tegnet efter '(').
 */
function findMatchingClosingDelim(str, startPos, openCh, closeCh) {
    let depth = 1;
    let i = startPos;
    while (i < str.length && depth > 0) {
        if (str[i] === '{' && str[i - 1] !== '\\') {
            // Spring over indre LaTeX-grupper så parenteser indeni ikke forstyrrer
            const grp = readLatexGroupOrToken(str, i);
            i = grp.nextPos;
            continue;
        }
        if (str[i] === openCh) depth++;
        else if (str[i] === closeCh) depth--;
        i++;
    }
    return depth === 0 ? i - 1 : -1;
}

/**
 * Kompilerer en LaTeX-streng rekursivt til en sekvens af OMML-noder (<m:r>, <m:f>, <m:rad>, <m:sSup>, <m:d> osv.).
 */
function compileLatexFragmentToOmml(latex) {
    const str = sanitizeLatex(latex);
    const nodes = [];
    let i = 0;

    // Hjælper til at vedhæfte eksponent/subscript på det seneste element
    const attachSubSup = (baseOmml, subOmml, supOmml) => {
        const safeBase = baseOmml || makeRun(' ');
        if (subOmml !== null && supOmml !== null) {
            return `<m:sSubSup><m:e>${safeBase}</m:e><m:sub>${subOmml}</m:sub><m:sup>${supOmml}</m:sup></m:sSubSup>`;
        }
        if (supOmml !== null) {
            return `<m:sSup><m:e>${safeBase}</m:e><m:sup>${supOmml}</m:sup></m:sSup>`;
        }
        return `<m:sSub><m:e>${safeBase}</m:e><m:sub>${subOmml}</m:sub></m:sSub>`;
    };

    while (i < str.length) {
        const ch = str[i];

        // 1. Mellemrum
        if (ch === ' ') {
            i++;
            continue;
        }

        // 2. LaTeX-kommandoer (\frac, \sqrt, \sin, \pi osv.)
        if (ch === '\\') {
            let cmdEnd = i + 1;
            while (cmdEnd < str.length && /[a-zA-Z]/.test(str[cmdEnd])) cmdEnd++;
            if (cmdEnd === i + 1 && cmdEnd < str.length) cmdEnd++;
            const cmd = str.substring(i, cmdEnd);
            i = cmdEnd;

            // \frac{num}{den}
            if (cmd === '\\frac') {
                const numGrp = readLatexGroupOrToken(str, i);
                i = numGrp.nextPos;
                const denGrp = readLatexGroupOrToken(str, i);
                i = denGrp.nextPos;
                const numOmml = compileLatexFragmentToOmml(numGrp.content) || makeRun(' ');
                const denOmml = compileLatexFragmentToOmml(denGrp.content) || makeRun(' ');
                nodes.push(`<m:f><m:fPr><m:type m:val="bar"/></m:fPr><m:num>${numOmml}</m:num><m:den>${denOmml}</m:den></m:f>`);
                continue;
            }

            // \sqrt[deg]{arg} eller \sqrt{arg}
            if (cmd === '\\sqrt') {
                while (i < str.length && str[i] === ' ') i++;
                let degContent = null;
                if (i < str.length && str[i] === '[') {
                    const closeBracket = str.indexOf(']', i + 1);
                    if (closeBracket !== -1) {
                        degContent = str.substring(i + 1, closeBracket);
                        i = closeBracket + 1;
                    }
                }
                const argGrp = readLatexGroupOrToken(str, i);
                i = argGrp.nextPos;
                const eOmml = compileLatexFragmentToOmml(argGrp.content) || makeRun(' ');
                if (degContent) {
                    const degOmml = compileLatexFragmentToOmml(degContent) || makeRun(' ');
                    nodes.push(`<m:rad><m:radPr><m:degHide m:val="0"/></m:radPr><m:deg>${degOmml}</m:deg><m:e>${eOmml}</m:e></m:rad>`);
                } else {
                    nodes.push(`<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e>${eOmml}</m:e></m:rad>`);
                }
                continue;
            }

            // \text{...} eller \mathrm{...}
            if (cmd === '\\text' || cmd === '\\mathrm') {
                const txtGrp = readLatexGroupOrToken(str, i);
                i = txtGrp.nextPos;
                nodes.push(makeRun(txtGrp.content, true));
                continue;
            }

            // Kendte symboler (\pi, \Rightarrow, \approx osv.)
            if (Object.prototype.hasOwnProperty.call(LATEX_SYMBOLS, cmd)) {
                nodes.push(makeRun(LATEX_SYMBOLS[cmd], false));
                continue;
            }

            // Matematiske funktioner (\sin, \cos, \ln osv.)
            const bareCmd = cmd.slice(1);
            if (UPRIGHT_FUNCTIONS.has(bareCmd)) {
                nodes.push(makeRun(bareCmd, true));
                continue;
            }

            // Ukendt kommando: udskriv selve navnet
            nodes.push(makeRun(bareCmd, false));
            continue;
        }

        // 3. Eksplicit krøllet parentes-gruppe { ... }
        if (ch === '{') {
            const grp = readLatexGroupOrToken(str, i);
            i = grp.nextPos;
            const innerOmml = compileLatexFragmentToOmml(grp.content);
            if (innerOmml) nodes.push(innerOmml);
            continue;
        }

        // 4. Potens (^) og Subscript (_)
        if (ch === '^' || ch === '_') {
            let subOmml = null;
            let supOmml = null;

            while (i < str.length && (str[i] === '^' || str[i] === '_')) {
                const kind = str[i];
                i++;
                const grp = readLatexGroupOrToken(str, i);
                i = grp.nextPos;
                const compiled = compileLatexFragmentToOmml(grp.content) || makeRun(' ');
                if (kind === '^') supOmml = compiled;
                else subOmml = compiled;
                while (i < str.length && str[i] === ' ') i++;
            }

            const baseNode = nodes.length > 0 ? nodes.pop() : makeRun(' ');
            nodes.push(attachSubSup(baseNode, subOmml, supOmml));
            continue;
        }

        // 5. Skalerende parenteser (...)
        if (ch === '(') {
            const closeIdx = findMatchingClosingDelim(str, i + 1, '(', ')');
            if (closeIdx !== -1) {
                const innerLatex = str.substring(i + 1, closeIdx);
                i = closeIdx + 1;
                const innerOmml = compileLatexFragmentToOmml(innerLatex) || makeRun(' ');
                nodes.push(`<m:d><m:dPr><m:begChr m:val="("/><m:endChr m:val=")"/></m:dPr><m:e>${innerOmml}</m:e></m:d>`);
                continue;
            }
        }

        // 6. Sammenhængende tal (inkl. dansk decimalkomma "3,14" eller punktum "3.14")
        if (/\d/.test(ch)) {
            let numEnd = i + 1;
            while (numEnd < str.length) {
                if (/\d/.test(str[numEnd])) {
                    numEnd++;
                } else if ((str[numEnd] === ',' || str[numEnd] === '.') && numEnd + 1 < str.length && /\d/.test(str[numEnd + 1])) {
                    numEnd += 2;
                } else {
                    break;
                }
            }
            nodes.push(makeRun(str.substring(i, numEnd), false));
            i = numEnd;
            continue;
        }

        // 7. Operatorer med pæn typografisk luft
        if (ch === '+' || ch === '=' || ch === '<' || ch === '>') {
            nodes.push(makeRun(` ${ch} `, false));
            i++;
            continue;
        }
        if (ch === '-') {
            nodes.push(makeRun('−', false));
            i++;
            continue;
        }

        // 8. Enkeltbogstaver (variabler som x, y, a, b, C) og øvrige tegn
        nodes.push(makeRun(ch, false));
        i++;
    }

    return nodes.join('');
}

/**
 * Konverterer en LaTeX-streng til et komplet <m:oMath>-element.
 */
window.latexToOmmlString = function(latexStr) {
    const inner = compileLatexFragmentToOmml(latexStr || "");
    return `<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">${inner}</m:oMath>`;
};

/**
 * Pakker et OMML-fragment i en Flat OPC Word-pakke.
 */
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

window.latexToOoxml = function(latexStr) {
    const omml = window.latexToOmmlString(latexStr);
    return buildFlatOpc(omml);
};

// ─────────────────────────────────────────────────────────────────
//  DEL 2: OMML -> Matematisk Udtryk Læser (Læser rigtige Word-ligninger!)
// ─────────────────────────────────────────────────────────────────

function getLocalName(node) {
    if (!node) return "";
    if (node.localName) return node.localName;
    const name = node.nodeName || "";
    const colonIdx = name.indexOf(':');
    return colonIdx !== -1 ? name.substring(colonIdx + 1) : name;
}

function findDirectChild(parent, localName) {
    if (!parent || !parent.childNodes) return null;
    for (let i = 0; i < parent.childNodes.length; i++) {
        const child = parent.childNodes[i];
        if (getLocalName(child) === localName) return child;
    }
    return null;
}

function wrapIfCompound(expr) {
    const trimmed = String(expr || "").trim();
    if (!trimmed) return "";
    if (/^[a-zA-Z0-9.,πΠ]+$/.test(trimmed)) return trimmed;
    if (isFullyWrappedInParens(trimmed)) return trimmed;
    return `(${trimmed})`;
}

/**
 * Går rekursivt igennem en OMML DOM-node (<m:oMath>, <m:f>, <m:sSup>, <m:rad>, <m:d> osv.)
 * og genskaber et præcist matematisk tekst-udtryk (f.eks. "(x+1)/(x-1) = 2").
 */
function parseOmmlDomNode(node) {
    if (!node) return "";
    if (node.nodeType === 3) {
        return node.nodeValue || "";
    }
    if (node.nodeType !== 1) return "";

    const local = getLocalName(node);

    // Ignorér egenskabs-noder (rPr, fPr, radPr, dPr, ctrlPr osv.)
    if (local.endsWith("Pr") || local === "argPr") {
        return "";
    }

    // Brøk: <m:f><m:num>...</m:num><m:den>...</m:den></m:f>
    if (local === "f") {
        const numNode = findDirectChild(node, "num");
        const denNode = findDirectChild(node, "den");
        const num = wrapIfCompound(parseOmmlChildren(numNode));
        const den = wrapIfCompound(parseOmmlChildren(denNode));
        return `${num}/${den}`;
    }

    // Potens: <m:sSup><m:e>...</m:e><m:sup>...</m:sup></m:sSup>
    if (local === "sSup") {
        const eNode = findDirectChild(node, "e");
        const supNode = findDirectChild(node, "sup");
        const base = wrapIfCompound(parseOmmlChildren(eNode));
        const supRaw = parseOmmlChildren(supNode).trim();
        const sup = /^[a-zA-Z0-9]+$/.test(supRaw) ? supRaw : `(${supRaw})`;
        return `${base}^${sup}`;
    }

    // Subscript: <m:sSub><m:e>...</m:e><m:sub>...</m:sub></m:sSub>
    if (local === "sSub") {
        const eNode = findDirectChild(node, "e");
        const subNode = findDirectChild(node, "sub");
        return `${parseOmmlChildren(eNode)}_${parseOmmlChildren(subNode).trim()}`;
    }

    // Sub- og Superscript: <m:sSubSup>
    if (local === "sSubSup") {
        const eNode = findDirectChild(node, "e");
        const supNode = findDirectChild(node, "sup");
        const base = wrapIfCompound(parseOmmlChildren(eNode));
        const sup = wrapIfCompound(parseOmmlChildren(supNode));
        return `${base}^${sup}`;
    }

    // Kvadratrod / n'te rod: <m:rad><m:deg>...</m:deg><m:e>...</m:e></m:rad>
    if (local === "rad") {
        const degNode = findDirectChild(node, "deg");
        const eNode = findDirectChild(node, "e");
        const deg = degNode ? parseOmmlChildren(degNode).trim() : "";
        const inner = parseOmmlChildren(eNode).trim();
        if (deg && deg !== "2") {
            return `(${inner})^(1/(${deg}))`;
        }
        return `sqrt(${inner})`;
    }

    // Parenteser / Delimiters: <m:d><m:dPr>...</m:dPr><m:e>...</m:e></m:d>
    if (local === "d") {
        const dPr = findDirectChild(node, "dPr");
        let begChr = "(";
        let endChr = ")";
        if (dPr) {
            const begNode = findDirectChild(dPr, "begChr");
            const endNode = findDirectChild(dPr, "endChr");
            if (begNode) begChr = begNode.getAttribute("m:val") || begNode.getAttribute("val") || "(";
            if (endNode) endChr = endNode.getAttribute("m:val") || endNode.getAttribute("val") || ")";
        }
        const parts = [];
        for (let i = 0; i < node.childNodes.length; i++) {
            if (getLocalName(node.childNodes[i]) === "e") {
                parts.push(parseOmmlChildren(node.childNodes[i]));
            }
        }
        const inner = parts.join(", ");
        if (begChr === "|" && endChr === "|") {
            return `abs(${inner})`;
        }
        return `(${inner})`;
    }

    // Funktion: <m:func><m:fName>sin</m:fName><m:e>x</m:e></m:func>
    if (local === "func") {
        const fNameNode = findDirectChild(node, "fName");
        const eNode = findDirectChild(node, "e");
        const fnName = parseOmmlChildren(fNameNode).trim();
        const arg = parseOmmlChildren(eNode).trim();
        if (isFullyWrappedInParens(arg)) {
            return `${fnName}${arg}`;
        }
        return `${fnName}(${arg})`;
    }

    // Tekst-node (<m:t> eller <w:t>)
    if (local === "t") {
        return node.textContent || "";
    }

    return parseOmmlChildren(node);
}

function parseOmmlChildren(parent) {
    if (!parent || !parent.childNodes) return "";
    let out = "";
    for (let i = 0; i < parent.childNodes.length; i++) {
        out += parseOmmlDomNode(parent.childNodes[i]);
    }
    return out;
}

/**
 * Udtrækker et matematisk udtryk fra en OOXML-streng (f.eks. fra selection.getOoxml()).
 * Hvis OOXML'en indeholder et eller flere <m:oMath>-elementer, konverteres de til
 * korrekt matematisk syntaks med brøker, potenser og rødder bevaret!
 * Returnerer null hvis der ingen <m:oMath> er i OOXML'en.
 */
window.ommlToExpression = function(ooxmlString) {
    if (!ooxmlString || typeof ooxmlString !== "string") return null;
    if (!ooxmlString.includes("oMath")) return null;

    try {
        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(ooxmlString, "text/xml");
        if (xmlDoc.querySelector("parsererror")) return null;

        const allNodes = xmlDoc.getElementsByTagName("*");
        const oMathNodes = [];
        for (let i = 0; i < allNodes.length; i++) {
            if (getLocalName(allNodes[i]) === "oMath") {
                oMathNodes.push(allNodes[i]);
            }
        }
        if (oMathNodes.length === 0) return null;

        const expressions = oMathNodes
            .map(node => parseOmmlDomNode(node).trim())
            .filter(Boolean);

        if (expressions.length === 0) return null;

        let combined = expressions.join(" ").replace(/\s+/g, " ").trim();

        // Hvis ligningen allerede indeholder en tidligere pil-løsning (f.eks. "2x+5=15 ⇒ x=5"),
        // behold selve opgave-udtrykket før pilen, så eleven kan regne videre på det!
        if (combined.includes("⇒")) {
            combined = combined.split("⇒")[0].trim();
        }
        // Tilføj parentes omkring funktionsargumenter hvis Word havde "sin x" uden parentes
        combined = combined.replace(/\b(sin|cos|tan|ln|log|sqrt)\s+([a-zA-Z0-9]+)\b/g, '$1($2)');

        return combined || null;
    } catch (e) {
        console.warn("Kunne ikke parse OMML fra Word:", e);
        return null;
    }
};

// ─────────────────────────────────────────────────────────────────
//  DEL 3: KaTeX HTML-rendering & Brøk-formatering
// ─────────────────────────────────────────────────────────────────

window.renderLatexToHtml = function(latexStr, displayMode) {
    if (!latexStr) return "";
    if (typeof window.katex !== 'undefined') {
        try {
            return window.katex.renderToString(sanitizeLatex(latexStr), {
                throwOnError: false,
                displayMode: Boolean(displayMode)
            });
        } catch (e) {
            // Fald tilbage nedenfor
        }
    }
    return escapeXml(latexStr);
};

window.formatFractions = function(fullLatex, isStacked) {
    if (!fullLatex || typeof fullLatex !== 'string') return "";

    if (isStacked) {
        let result = fullLatex.replace(/\s*\/\s*/g, '/');
        let index = result.indexOf('/');
        let safetyCounter = 0;

        while (index !== -1 && safetyCounter < 100) {
            safetyCounter++;
            let leftStart = index - 1;
            let num = '';
            if (leftStart >= 0 && result[leftStart] === ')') {
                let depth = 1;
                leftStart--;
                while (leftStart >= 0 && depth > 0) {
                    if (result[leftStart] === ')') depth++;
                    else if (result[leftStart] === '(') depth--;
                    leftStart--;
                }
                leftStart++;
                num = result.substring(leftStart + 1, index - 1);
            } else {
                while (leftStart >= 0) {
                    if (result[leftStart] === '}') {
                        let braceDepth = 1;
                        leftStart--;
                        while (leftStart >= 0 && braceDepth > 0) {
                            if (result[leftStart] === '}') braceDepth++;
                            else if (result[leftStart] === '{') braceDepth--;
                            leftStart--;
                        }
                    } else if (/[a-zA-Z0-9_.,\\^]/.test(result[leftStart])) {
                        leftStart--;
                    } else {
                        break;
                    }
                }
                leftStart++;
                num = result.substring(leftStart, index);
            }

            let rightEnd = index + 1;
            let den = '';
            if (rightEnd < result.length && result[rightEnd] === '(') {
                let depth = 1;
                rightEnd++;
                while (rightEnd < result.length && depth > 0) {
                    if (result[rightEnd] === '(') depth++;
                    else if (result[rightEnd] === ')') depth--;
                    rightEnd++;
                }
                rightEnd--;
                den = result.substring(index + 2, rightEnd);
            } else {
                while (rightEnd < result.length) {
                    if (result[rightEnd] === '{') {
                        let braceDepth = 1;
                        rightEnd++;
                        while (rightEnd < result.length && braceDepth > 0) {
                            if (result[rightEnd] === '{') braceDepth++;
                            else if (result[rightEnd] === '}') braceDepth--;
                            rightEnd++;
                        }
                    } else if (/[a-zA-Z0-9_.,\\^]/.test(result[rightEnd])) {
                        rightEnd++;
                    } else {
                        break;
                    }
                }
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
        let result = fullLatex;
        let index = result.indexOf('\\frac{');
        let safetyCounter = 0;

        while (index !== -1 && safetyCounter < 100) {
            safetyCounter++;
            let p1Start = index + 5;
            if (result[p1Start] !== '{') {
                index = result.indexOf('\\frac{', index + 6);
                continue;
            }
            let depth1 = 1, p1End = p1Start + 1;
            while (p1End < result.length && depth1 > 0) {
                if (result[p1End] === '{' && result[p1End - 1] !== '\\') depth1++;
                else if (result[p1End] === '}' && result[p1End - 1] !== '\\') depth1--;
                p1End++;
            }
            p1End--;
            let num = result.substring(p1Start + 1, p1End);

            let p2Start = p1End + 1;
            while (p2Start < result.length && result[p2Start] === ' ') p2Start++;
            if (result[p2Start] !== '{') {
                index = result.indexOf('\\frac{', p1End + 1);
                continue;
            }
            let depth2 = 1, p2End = p2Start + 1;
            while (p2End < result.length && depth2 > 0) {
                if (result[p2End] === '{' && result[p2End - 1] !== '\\') depth2++;
                else if (result[p2End] === '}' && result[p2End - 1] !== '\\') depth2--;
                p2End++;
            }
            p2End--;
            let den = result.substring(p2Start + 1, p2End);

            if (/[\+\-\*\=\s]/.test(num) && !isFullyWrappedInParens(num)) num = `(${num})`;
            if (/[\+\-\*\=\s]/.test(den) && !isFullyWrappedInParens(den)) den = `(${den})`;

            result = result.substring(0, index) + `${num}/${den}` + result.substring(p2End + 1);
            index = result.indexOf('\\frac{');
        }
        return result;
    }
};
