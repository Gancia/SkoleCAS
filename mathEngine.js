// ============================================================================
// SkoleCAS Kerne:
//  1. window.CAS           – Udskiftelig CAS-adapter (eval, solve, diff, int)
//  2. window.WordBridge    – To-vejs bro til Word (læser OG skriver <m:oMath> OMML)
//  3. window.GeoGebraBridge– Lazy-loaded GeoGebra-graftegner & ekstern bro
// ============================================================================

window.calculationHistory = [];
window.lastCalculation = null;
let panelInputDirty = false;
let isCalculatingInline = false;

window.clearPanelCalculationState = function() {
    window.lastCalculation = null;
    panelInputDirty = false;
};

// ─────────────────────────────────────────────────────────────────
//  1. CAS-ADAPTER (Isoleret beregningslag)
// ─────────────────────────────────────────────────────────────────

function detectMainVariable(mathExpr) {
    try {
        const exprForVars = mathExpr.includes('=')
            ? '(' + mathExpr.split('=')[0] + ')-(' + mathExpr.split('=').slice(1).join('=') + ')'
            : mathExpr;
        const vars = nerdamer(exprForVars).variables();
        if (Array.isArray(vars) && vars.length > 0) {
            if (vars.includes('x')) return 'x';
            const nonE = vars.filter(v => v !== 'e');
            return nonE.length > 0 ? nonE[0] : vars[0];
        }
    } catch (e) {
        // Fallback nedenfor
    }

    const fnNames = [
        'arcsin', 'arccos', 'arctan', 'sinh', 'cosh', 'tanh',
        'asin', 'acos', 'atan', 'sqrt', 'sin', 'cos', 'tan',
        'cot', 'sec', 'csc', 'log', 'exp', 'abs', 'pi', 'ln', 'e'
    ];
    let cleaned = mathExpr;
    for (const fn of fnNames) {
        cleaned = cleaned.replace(new RegExp('\\b' + fn + '\\b', 'gi'), '');
    }
    const matches = cleaned.match(/[a-zA-Z]/g) || [];
    if (matches.includes('x')) return 'x';
    return matches.length > 0 ? matches[0] : 'x';
}

function roundNumbersInString(str, decimals) {
    let wasRounded = false;
    const roundedStr = str.replace(/\d+\.\d+/g, function(match) {
        const originalNum = parseFloat(match);
        if (!Number.isFinite(originalNum)) return match;
        const roundedNum = window.safeRound(originalNum, decimals);
        if (Math.abs(originalNum - roundedNum) > 1e-12 || /\.\d{7,}/.test(match)) {
            wasRounded = true;
        }
        let formatted = roundedNum.toFixed(decimals)
            .replace(/(\.\d*[1-9])0+$/, '$1')
            .replace(/\.0*$/, '');
        if (formatted === '-0') formatted = '0';
        return formatted;
    });
    return { text: roundedStr, wasRounded };
}

window.CAS = {
    detectVariable: detectMainVariable,

    /**
     * Konverterer et råt udtryk til GeoGebra-kompatibel syntaks.
     */
    toGeoGebraSyntax: function(rawInput) {
        if (!rawInput) return "";
        let s = String(rawInput).trim();
        if (s.includes('⇒')) s = s.split('⇒')[0].trim();
        if (s.includes('≈')) s = s.split('≈')[0].trim();
        s = window.normalizeInput(s);
        s = window.convertDanishDecimalToDot(s, ",");
        if (s.endsWith('=')) s = s.slice(0, -1).trim();
        return s;
    },

    /**
     * Formaterer et råt matematisk tekstsegment strukturelt til LaTeX
     * uden at Nerdamer forhåndsberegner/reducerer udtrykket (så f.eks. 1/2 + 3/4 forbliver \frac{1}{2} + \frac{3}{4}).
     */
    formatSegmentToLatex: function(segment, decSep, isStacked) {
        if (!segment || !segment.trim()) return "";
        const sep = decSep || ",";
        const stacked = isStacked !== undefined ? isStacked : true;

        const supMap = {
            '⁰': '0', '¹': '1', '²': '2', '³': '3', '⁴': '4',
            '⁵': '5', '⁶': '6', '⁷': '7', '⁸': '8', '⁹': '9', '⁻': '-'
        };

        let s = segment.trim()
            .replace(/[\u2212\u2013\u2014]/g, '-')
            .replace(/[\u00B7\u22C5\u00D7]/g, '*')
            .replace(/[\u00F7:]/g, '/')
            .replace(/[πΠ]/g, '\\pi ')
            .replace(/(^|[^\\])\bpi\b/gi, '$1\\pi ')
            .replace(/[⁰¹²³⁴⁵⁶⁷⁸⁹⁻]+/g, (m) => {
                const digits = m.split('').map(ch => supMap[ch] || ch).join('');
                return `^{${digits}}`;
            })
            .replace(/√\s*\(/g, 'sqrt(')
            .replace(/√([a-zA-Z0-9.,]+)/g, '\\sqrt{$1}');

        // Konvertér balancerede sqrt(...) til \sqrt{...}
        let sqrtMatch = /\bsqrt\s*\(/i.exec(s);
        let guard = 0;
        while (sqrtMatch && guard < 50) {
            guard++;
            const startIdx = sqrtMatch.index;
            const openParenIdx = startIdx + sqrtMatch[0].length - 1;
            let depth = 1;
            let i = openParenIdx + 1;
            while (i < s.length && depth > 0) {
                if (s[i] === '(') depth++;
                else if (s[i] === ')') depth--;
                i++;
            }
            if (depth === 0) {
                const inner = s.substring(openParenIdx + 1, i - 1);
                s = s.substring(0, startIdx) + `\\sqrt{${inner}}` + s.substring(i);
            } else {
                break;
            }
            sqrtMatch = /\bsqrt\s*\(/i.exec(s);
        }

        // Konvertér balancerede potenser ^(...) til ^{...}
        let powMatch = /\^\s*\(/.exec(s);
        guard = 0;
        while (powMatch && guard < 50) {
            guard++;
            const startIdx = powMatch.index;
            const openParenIdx = startIdx + powMatch[0].length - 1;
            let depth = 1;
            let i = openParenIdx + 1;
            while (i < s.length && depth > 0) {
                if (s[i] === '(') depth++;
                else if (s[i] === ')') depth--;
                i++;
            }
            if (depth === 0) {
                const inner = s.substring(openParenIdx + 1, i - 1);
                s = s.substring(0, startIdx) + `^{${inner}}` + s.substring(i);
            } else {
                break;
            }
            powMatch = /\^\s*\(/.exec(s);
        }

        // Konvertér simple potenser ^2, ^-2, ^x til ^{...}
        s = s.replace(/\^\s*(-?[a-zA-Z0-9.,]+)/g, '^{$1}');

        // Konvertér kendte funktioner til LaTeX-kommandoer
        s = s.replace(/(^|[^\\])\b(arcsin|arccos|arctan|asin|acos|atan|sin|cos|tan|ln|log)\s*\(/gi, (m, prefix, fn) => {
            return `${prefix}\\${fn.toLowerCase()}(`;
        });

        // Formatér brøker (1/2 -> \frac{1}{2}) uden at evaluere talværdierne
        if (window.formatFractions) {
            s = window.formatFractions(s, stacked);
        }

        // Erstat eksplicit multiplikation (*) med \cdot
        s = s.replace(/\s*\*\s*/g, ' \\cdot ');

        // Tilpas decimalseparator til brugerens valg
        if (sep === ",") {
            s = s.replace(/(\d)\.(\d)/g, '$1,$2');
        } else {
            s = s.replace(/(\d),(\d)/g, '$1.$2');
        }

        return s.replace(/\s{2,}/g, ' ').trim();
    },

    /**
     * Konverterer et brugerudtryk til LaTeX til forhåndsvisning og venstreside af udregninger.
     */
    toPreviewLatex: function(rawInput, decSep, isStacked) {
        if (!rawInput || !rawInput.trim()) return "";
        let cleaned = rawInput.trim();
        if (cleaned.includes('⇒')) cleaned = cleaned.split('⇒')[0].trim();
        if (cleaned.includes('≈')) cleaned = cleaned.split('≈')[0].trim();

        const sep = decSep || ",";
        const stacked = isStacked !== undefined ? isStacked : true;

        if (cleaned.includes('=')) {
            const parts = cleaned.split('=');
            const leftRaw = parts[0].trim();
            const rightRaw = parts.slice(1).join('=').trim();
            const leftTex = this.formatSegmentToLatex(leftRaw, sep, stacked);
            const rightTex = rightRaw ? this.formatSegmentToLatex(rightRaw, sep, stacked) : "";
            return rightRaw ? `${leftTex} = ${rightTex}` : `${leftTex} =`;
        }

        return this.formatSegmentToLatex(cleaned, sep, stacked);
    },

    /**
     * Udfører CAS-beregning ('eval', 'solve', 'diff', 'int').
     */
    execute: function(rawText, actionType, options) {
        const opts = options || {};
        const isExact = Boolean(opts.isExact);
        const decimals = Number.isFinite(opts.decimals) ? opts.decimals : 2;
        const decSep = opts.decSep || ",";
        const isStacked = opts.isStacked !== undefined ? Boolean(opts.isStacked) : true;

        let textToParse = String(rawText || "").trim();
        if (!textToParse) {
            throw new Error("Skriv et udtryk i feltet eller markér en ligning i Word først.");
        }

        if (textToParse.includes('⇒')) {
            textToParse = textToParse.split('⇒')[0].trim();
        }
        if (textToParse.includes('≈')) {
            textToParse = textToParse.split('≈')[0].trim();
        }

        const isFnOrYDef = /^([a-zA-Z]\s*\(\s*[a-zA-Z]\s*\)|y)\s*=/i.test(textToParse);
        if ((actionType === 'eval' || actionType === 'approx' || actionType === 'calc' || actionType === 'diff' || actionType === 'int') && !isFnOrYDef && textToParse.includes('=')) {
            textToParse = textToParse.split('=')[0].trim();
        } else if (textToParse.endsWith('=')) {
            textToParse = textToParse.slice(0, -1).trim();
        }

        let mathExpr = window.normalizeInput(textToParse);
        mathExpr = window.convertDanishDecimalToDot(mathExpr, decSep);

        const fnHeaderMatch = mathExpr.match(/^([a-zA-Z])\s*\(\s*([a-zA-Z])\s*\)\s*=\s*(.+)$/);
        const yHeaderMatch = !fnHeaderMatch ? mathExpr.match(/^y\s*=\s*(.+)$/i) : null;

        let originalLatex = "";
        try {
            originalLatex = window.CAS.toPreviewLatex(textToParse, decSep, isStacked);
        } catch (e) {
            originalLatex = textToParse;
        }

        const mainVar = fnHeaderMatch ? fnHeaderMatch[2] : detectMainVariable(mathExpr);
        let latexResult = "";
        let prefix = "";

        try {
            if (actionType === 'solve') {
                let eqToSolve = mathExpr;
                if (fnHeaderMatch) eqToSolve = fnHeaderMatch[3] + "=0";
                else if (yHeaderMatch) eqToSolve = yHeaderMatch[1] + "=0";
                else if (!eqToSolve.includes('=')) eqToSolve = eqToSolve + "=0";

                const parsed = nerdamer.solveEquations(eqToSolve, mainVar);
                const rawSolutions = Array.isArray(parsed) ? parsed : [parsed];

                const realSolutions = rawSolutions.filter(sol => {
                    const s = sol ? sol.toString() : "";
                    return s !== "" && !/\bi\b|\*i/.test(s);
                });

                if (realSolutions.length === 0) {
                    throw new Error("Ligningen har ingen reelle løsninger.");
                }

                let anyApprox = false;
                const formattedSolutions = realSolutions.map(val => {
                    if (isExact) {
                        return nerdamer(val).toTeX();
                    } else {
                        const decText = nerdamer(val).evaluate().text('decimals');
                        const { text: roundedText, wasRounded } = roundNumbersInString(decText, decimals);
                        if (wasRounded) anyApprox = true;
                        return nerdamer(roundedText).toTeX('decimals');
                    }
                });

                const joinSymbol = anyApprox ? ` \\lor ${mainVar} \\approx ` : ` \\lor ${mainVar} = `;
                latexResult = formattedSolutions.join(joinSymbol);
                prefix = anyApprox ? `\\Rightarrow ${mainVar} \\approx ` : `\\Rightarrow ${mainVar} = `;
            }
            else if (actionType === 'eval' || actionType === 'approx' || actionType === 'calc') {
                let exprToEval = mathExpr;
                if (fnHeaderMatch) exprToEval = fnHeaderMatch[3];
                else if (yHeaderMatch) exprToEval = yHeaderMatch[1];
                else if (exprToEval.endsWith('=')) exprToEval = exprToEval.slice(0, -1).trim();

                if (isExact) {
                    latexResult = nerdamer(exprToEval).evaluate().toTeX();
                    prefix = "= ";
                } else {
                    const decimalStr = nerdamer(exprToEval).evaluate().text('decimals');
                    const { text: roundedStr, wasRounded } = roundNumbersInString(decimalStr, decimals);
                    latexResult = nerdamer(roundedStr).toTeX('decimals');
                    prefix = wasRounded ? "\\approx " : "= ";
                }
            }
            else if (actionType === 'diff') {
                let exprToDiff = fnHeaderMatch ? fnHeaderMatch[3] : (yHeaderMatch ? yHeaderMatch[1] : mathExpr);
                if (exprToDiff.includes('=')) exprToDiff = exprToDiff.split('=')[0].trim();
                const fnName = fnHeaderMatch ? fnHeaderMatch[1] : 'f';
                const diffEvaluated = nerdamer(`diff(${exprToDiff}, ${mainVar})`).evaluate();
                if (isExact) {
                    latexResult = diffEvaluated.toTeX();
                } else {
                    const decimalStr = diffEvaluated.text('decimals');
                    const { text: roundedStr } = roundNumbersInString(decimalStr, decimals);
                    latexResult = nerdamer(roundedStr).toTeX('decimals');
                }
                prefix = `\\Rightarrow ${fnName}'(${mainVar}) = `;
            }
            else if (actionType === 'int') {
                let exprToInt = fnHeaderMatch ? fnHeaderMatch[3] : (yHeaderMatch ? yHeaderMatch[1] : mathExpr);
                if (exprToInt.includes('=')) exprToInt = exprToInt.split('=')[0].trim();
                const fnName = fnHeaderMatch ? fnHeaderMatch[1] : 'f';
                const intEvaluated = nerdamer(`integrate(${exprToInt}, ${mainVar})`);
                if (isExact) {
                    latexResult = intEvaluated.toTeX() + " + C";
                } else {
                    const decimalStr = intEvaluated.text('decimals');
                    const { text: roundedStr } = roundNumbersInString(decimalStr, decimals);
                    latexResult = nerdamer(roundedStr).toTeX('decimals') + " + C";
                }
                prefix = `\\Rightarrow \\int ${fnName}(${mainVar})\\,d${mainVar} = `;
            }
            else {
                throw new Error("Ukendt beregningstype: " + actionType);
            }
        } catch (calcError) {
            if (calcError.message === "Ligningen har ingen reelle løsninger.") {
                throw calcError;
            }
            throw new Error(`Kunne ikke beregne udtrykket "${textToParse}". Kontrollér parenteser og syntaks.`);
        }

        if (window.formatFractions) {
            latexResult = window.formatFractions(latexResult, isStacked);
            originalLatex = window.formatFractions(originalLatex, isStacked);
        }

        if (decSep === "," && latexResult) {
            latexResult = latexResult.replace(/(\d)\.(\d)/g, '$1,$2');
        }

        if (latexResult) {
            latexResult = latexResult.replace(/\\cdot\s*\\pi/g, '\\pi');
        }

        const answerOnlyLatex = (prefix + latexResult).trim();
        const fullLatex = `${originalLatex} ${answerOnlyLatex}`.trim();

        const toReadableText = (tex) => {
            return tex
                .replace(/\\approx\s*/g, '≈ ')
                .replace(/\\Rightarrow\s*/g, '⇒ ')
                .replace(/\\int\s*/g, '∫ ')
                .replace(/\\,/g, ' ')
                .replace(/\\pi/g, 'π')
                .replace(/\\cdot/g, '⋅')
                .replace(/\\lor/g, ' ∨ ')
                .replace(/\\sqrt\{([^}]+)\}/g, '√($1)')
                .replace(/\\frac\{([^}]+)\}\{([^}]+)\}/g, '($1)/($2)')
                .replace(/\^\{([^}]+)\}/g, '^$1')
                .replace(/\\/g, '')
                .replace(/\s+/g, ' ')
                .trim();
        };

        const fallbackText = toReadableText(answerOnlyLatex);
        const fullFallbackText = `${textToParse} ${fallbackText}`.trim();

        return {
            rawInput: textToParse,
            actionType,
            originalLatex,
            prefix,
            resultLatex: latexResult,
            answerOnlyLatex,
            fullLatex,
            fallbackText,
            fullFallbackText
        };
    }
};

// ─────────────────────────────────────────────────────────────────
//  2. WORDBRIDGE (To-vejs læsning & skrivning af OMML i Word)
// ─────────────────────────────────────────────────────────────────

window.WordBridge = {
    /**
     * Læser matematisk udtryk fra Word (både rigtige <m:oMath>-ligninger og rå tekst).
     * Returnerer { expr: string, sourceMode: 'word-highlight'|'word-line'|null, rawHighlightedText: string, hasColonPrefix: boolean }
     */
    readExpressionFromWord: async function(context) {
        const selection = context.document.getSelection();
        selection.load("text");
        const selOoxml = selection.getOoxml();
        await context.sync();

        // 1. Tjek først om brugeren har markeret en rigtig Word-ligning (<m:oMath>)
        if (window.ommlToExpression && selOoxml && selOoxml.value) {
            const ommlExpr = window.ommlToExpression(selOoxml.value);
            if (ommlExpr) {
                return {
                    expr: ommlExpr,
                    sourceMode: "word-highlight",
                    rawHighlightedText: selection.text ? selection.text.trim() : "",
                    hasColonPrefix: false
                };
            }
        }

        // 2. Tjek om brugeren har markeret almindelig tekst
        const highlightedText = selection.text ? selection.text.trim() : "";
        if (highlightedText) {
            return {
                expr: highlightedText,
                sourceMode: "word-highlight",
                rawHighlightedText: highlightedText,
                hasColonPrefix: false
            };
        }

        // 3. Ingen markering: Tjek aktuelt afsnit for enten tekst før markøren eller en <m:oMath>-ligning
        const paragraph = selection.paragraphs.getFirst();
        const rangeBeforeCursor = selection.expandTo(paragraph.getRange("Start"));
        rangeBeforeCursor.load("text");
        const paraOoxml = paragraph.getOoxml();
        await context.sync();

        const rawBeforeCursor = rangeBeforeCursor.text ? rangeBeforeCursor.text.trim() : "";
        const hasColonPrefix = rawBeforeCursor.includes(":");
        const textAfterColon = hasColonPrefix
            ? rawBeforeCursor.substring(rawBeforeCursor.lastIndexOf(":") + 1).trim()
            : rawBeforeCursor;

        if (window.ommlToExpression && paraOoxml && paraOoxml.value) {
            const paraOmmlExpr = window.ommlToExpression(paraOoxml.value);
            if (paraOmmlExpr) {
                return {
                    expr: paraOmmlExpr,
                    sourceMode: "word-line",
                    rawHighlightedText: "",
                    hasColonPrefix
                };
            }
        }

        if (textAfterColon) {
            return {
                expr: textAfterColon,
                sourceMode: "word-line",
                rawHighlightedText: "",
                hasColonPrefix
            };
        }

        return { expr: "", sourceMode: null, rawHighlightedText: "", hasColonPrefix: false };
    },

    /**
     * Indsætter et matematisk resultat som rigtig Word-ligning (<m:oMath>) uden låste bokse.
     */
    insertMathInWord: async function(context, _targetRange, latexToInsert, fallbackTextToInsert) {
        if (window.latexToOmmlString) {
            try {
                // Hvis markeringen allerede ligger inde i en gammel Content Control, pak den ud først
                try {
                    const preCheckSel = context.document.getSelection();
                    const existingParentCc = preCheckSel.parentContentControlOrNullObject;
                    existingParentCc.load("tag");
                    await context.sync();
                    if (!existingParentCc.isNullObject) {
                        existingParentCc.cannotDelete = false;
                        existingParentCc.cannotEdit = false;
                        existingParentCc.delete(true);
                        await context.sync();
                    }
                } catch (preErr) {
                    // Ignorér hvis der ikke er nogen eksisterende Content Control
                }

                const ommlString = window.latexToOmmlString(latexToInsert);
                const ccTag = "MATH_TARGET_" + Date.now();
                const activeRange = context.document.getSelection();
                const paragraph = activeRange.paragraphs.getFirst();

                const cc = activeRange.insertContentControl();
                cc.tag = ccTag;
                cc.cannotDelete = false;
                cc.cannotEdit = false;
                cc.insertText("[Beregner...]", "Replace");
                await context.sync();

                const pOoxml = paragraph.getOoxml();
                await context.sync();

                const parser = new DOMParser();
                const xmlDoc = parser.parseFromString(pOoxml.value, "text/xml");
                const mathDoc = parser.parseFromString(ommlString, "text/xml");
                const importedMathNode = xmlDoc.importNode(mathDoc.documentElement, true);

                const tags = xmlDoc.getElementsByTagName("*");
                let sdtNode = null;
                let rNode = null;

                for (let i = 0; i < tags.length; i++) {
                    if (tags[i].nodeName === "w:tag" || tags[i].localName === "tag") {
                        const val = tags[i].getAttribute("w:val") || tags[i].getAttribute("val");
                        if (val === ccTag) {
                            sdtNode = tags[i].parentNode.parentNode;
                            break;
                        }
                    }
                }

                if (!sdtNode) {
                    for (let i = 0; i < tags.length; i++) {
                        if ((tags[i].nodeName === "w:t" || tags[i].localName === "t") && tags[i].textContent.includes("[Beregner...]")) {
                            let p = tags[i].parentNode;
                            while (p) {
                                if (!rNode && (p.nodeName === "w:r" || p.localName === "r")) rNode = p;
                                if (!sdtNode && (p.nodeName === "w:sdt" || p.localName === "sdt")) sdtNode = p;
                                p = p.parentNode;
                            }
                            break;
                        }
                    }
                }

                let isInsideOMath = false;
                let checkParent = (sdtNode || rNode) ? (sdtNode || rNode).parentNode : null;
                while (checkParent) {
                    if (checkParent.nodeName === "m:oMath" || checkParent.localName === "oMath") {
                        isInsideOMath = true;
                        break;
                    }
                    checkParent = checkParent.parentNode;
                }

                const wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
                let nodeToInsert = xmlDoc.createDocumentFragment();
                if (isInsideOMath) {
                    while (importedMathNode.firstChild) {
                        nodeToInsert.appendChild(importedMathNode.firstChild);
                    }
                } else {
                    nodeToInsert.appendChild(importedMathNode);
                    // Tilføj et almindeligt tekst-run med mellemrum efter ligningen, så markøren kan komme ud af ligningsfeltet
                    const trailingRun = xmlDoc.createElementNS(wNs, "w:r");
                    const trailingText = xmlDoc.createElementNS(wNs, "w:t");
                    trailingText.setAttribute("xml:space", "preserve");
                    trailingText.textContent = " ";
                    trailingRun.appendChild(trailingText);
                    nodeToInsert.appendChild(trailingRun);
                }

                let replacedInXml = false;
                if (sdtNode) {
                    sdtNode.parentNode.replaceChild(nodeToInsert, sdtNode);
                    replacedInXml = true;
                } else if (rNode) {
                    rNode.parentNode.replaceChild(nodeToInsert, rNode);
                    replacedInXml = true;
                }

                if (replacedInXml) {
                    const serializer = new XMLSerializer();
                    let newOoxml = serializer.serializeToString(xmlDoc);

                    const preambleMatch = pOoxml.value.match(/^(<\?xml[^>]+>\s*<\?mso-application[^>]+>\s*)/i);
                    const preamble = preambleMatch ? preambleMatch[1] : '<?xml version="1.0" standalone="yes"?>\n<?mso-application progid="Word.Document"?>\n';
                    newOoxml = newOoxml.replace(/^<\?xml[^>]+>\s*/i, '');
                    newOoxml = preamble + newOoxml;

                    paragraph.insertOoxml(newOoxml, "Replace");
                    await context.sync();

                    // Lås op og fjern enhver tilbageværende midlertidig Content Control-ramme (bevar selve ligningen med delete(true))
                    try {
                        const sel = context.document.getSelection();
                        const parentCc = sel.parentContentControlOrNullObject;
                        parentCc.load("isNullObject");

                        const docCcs = context.document.contentControls;
                        docCcs.load("items/tag");
                        await context.sync();

                        if (!parentCc.isNullObject) {
                            parentCc.cannotDelete = false;
                            parentCc.cannotEdit = false;
                            parentCc.delete(true);
                        }

                        for (let i = 0; i < docCcs.items.length; i++) {
                            const item = docCcs.items[i];
                            if (item.tag && item.tag.startsWith("MATH_TARGET_")) {
                                item.cannotDelete = false;
                                item.cannotEdit = false;
                                item.delete(true);
                            }
                        }
                        await context.sync();

                        const afterRange = context.document.getSelection().getRange("End");
                        afterRange.select();
                        await context.sync();
                    } catch (unwrapErr) {
                        console.warn("Kunne ikke pakke Content Control ud efter indsættelse:", unwrapErr);
                    }
                    return;
                }
            } catch (ooxmlErr) {
                console.warn("OOXML-indsættelse fejlede, bruger tekst-fallback:", ooxmlErr);
            }
        }

        const fallbackRange = context.document.getSelection();
        fallbackRange.insertText(fallbackTextToInsert, "Replace");
        await context.sync();
    }
};

// ─────────────────────────────────────────────────────────────────
//  3. GEOGEBRA-BRO (Konverterer dansk matematik & åbner GeoGebra Online)
// ─────────────────────────────────────────────────────────────────

window.GeoGebraBridge = {
    /**
     * Henter det aktuelle udtryk (fra Word eller panelet), konverterer dansk komma
     * og potenser til gyldig GeoGebra-syntaks, kopierer det til udklipsholderen
     * og åbner GeoGebra Online i en ny browserfane.
     */
    openOnlineWithExpression: async function() {
        const noticeEl = document.getElementById("ggbCopyNotice");
        const inputEl = document.getElementById("mathInput");
        let expr = inputEl ? inputEl.value.trim() : "";

        // Hvis der er markeret en ligning i Word, hent den automatisk først
        if (window.isOfficeReady && window.Word) {
            try {
                await Word.run(async (context) => {
                    const wordData = await window.WordBridge.readExpressionFromWord(context);
                    if (wordData.sourceMode === "word-highlight" && wordData.expr) {
                        expr = wordData.expr;
                        if (inputEl) {
                            inputEl.value = expr;
                            window.updateMathPreview();
                        }
                    }
                });
            } catch (e) {
                // Fortsæt med panelets tekst
            }
        }

        const ggbExpr = expr ? window.CAS.toGeoGebraSyntax(expr) : "";

        if (ggbExpr && navigator.clipboard && navigator.clipboard.writeText) {
            try {
                await navigator.clipboard.writeText(ggbExpr);
                if (noticeEl) {
                    noticeEl.textContent = `Kopieret til GeoGebra: ${ggbExpr} (tryk Ctrl+V i GeoGebra)`;
                    noticeEl.classList.remove("hidden");
                }
            } catch (e) {
                if (noticeEl) {
                    noticeEl.textContent = `GeoGebra-format: ${ggbExpr}`;
                    noticeEl.classList.remove("hidden");
                }
            }
        } else if (ggbExpr && noticeEl) {
            noticeEl.textContent = `GeoGebra-format: ${ggbExpr}`;
            noticeEl.classList.remove("hidden");
        }

        window.open("https://www.geogebra.org/calculator", "_blank", "noopener,noreferrer");
    }
};

// ─────────────────────────────────────────────────────────────────
//  4. UI-INTEGRATION (Forhåndsvisning, Resultatkort & Historik)
// ─────────────────────────────────────────────────────────────────

window.getCalcOptions = function() {
    const exactToggle = document.getElementById("exactToggle");
    const stackedToggle = document.getElementById("stackedFractionToggle");
    const decimalSelect = document.getElementById("decimalSelect");
    const separatorSelect = document.getElementById("decimalSeparator");

    return {
        isExact: exactToggle ? exactToggle.checked : false,
        isStacked: stackedToggle ? stackedToggle.checked : true,
        decimals: decimalSelect ? parseInt(decimalSelect.value, 10) : 2,
        decSep: separatorSelect ? separatorSelect.value : ","
    };
};

window.updateMathPreview = function() {
    const inputEl = document.getElementById("mathInput");
    const previewEl = document.getElementById("mathPreview");
    if (!inputEl || !previewEl) return;

    const raw = inputEl.value.trim();
    if (!raw) {
        previewEl.innerHTML = '<span class="preview-placeholder">Forhåndsvisning vises her, mens du skriver...</span>';
        return;
    }

    const opts = window.getCalcOptions();
    const previewLatex = window.CAS.toPreviewLatex(raw, opts.decSep, opts.isStacked);
    if (window.renderLatexToHtml) {
        previewEl.innerHTML = window.renderLatexToHtml(previewLatex, false);
    } else {
        previewEl.textContent = raw;
    }
};

function renderResultInPanel(calcObj, skipHistoryPush) {
    window.lastCalculation = calcObj;

    const resultCard = document.getElementById("calcResultCard");
    const resultMathEl = document.getElementById("calcResultMath");
    const insertBtn = document.getElementById("insertResultToWordBtn");

    if (resultCard) resultCard.classList.remove("hidden");
    if (resultMathEl) {
        if (window.renderLatexToHtml) {
            resultMathEl.innerHTML = window.renderLatexToHtml(calcObj.fullLatex, true);
        } else {
            resultMathEl.textContent = calcObj.fullFallbackText;
        }
    }
    if (insertBtn) {
        insertBtn.disabled = false;
    }

    if (skipHistoryPush) {
        if (window.calculationHistory.length > 0 &&
            window.calculationHistory[0].rawInput === calcObj.rawInput &&
            window.calculationHistory[0].actionType === calcObj.actionType) {
            window.calculationHistory[0] = calcObj;
            renderHistoryList();
        }
        return;
    }

    if (
        window.calculationHistory.length === 0 ||
        window.calculationHistory[0].fullLatex !== calcObj.fullLatex
    ) {
        window.calculationHistory.unshift(calcObj);
        if (window.calculationHistory.length > 8) {
            window.calculationHistory.pop();
        }
        renderHistoryList();
    }
}

/**
 * Genberegner det aktive resultat i panelet øjeblikkeligt, når brugeren ændrer
 * indstillinger (Eksakt resultat, Opstillede brøker, Antal decimaler, Decimalseparator).
 */
window.refreshLastCalculationForSettings = function() {
    if (window.lastCalculation && window.lastCalculation.rawInput && window.lastCalculation.actionType) {
        try {
            const opts = window.getCalcOptions();
            const updatedCalc = window.CAS.execute(
                window.lastCalculation.rawInput,
                window.lastCalculation.actionType,
                opts
            );
            renderResultInPanel(updatedCalc, true);
        } catch (e) {
            // Ignorér hvis genberegning fejler
        }
    }
    if (window.refreshTriangleDisplay) {
        window.refreshTriangleDisplay();
    }
};

function renderHistoryList() {
    const container = document.getElementById("calcHistoryList");
    const section = document.getElementById("calcHistorySection");
    if (!container) return;

    if (window.calculationHistory.length === 0) {
        if (section) section.classList.add("hidden");
        container.innerHTML = "";
        return;
    }

    if (section) section.classList.remove("hidden");
    container.innerHTML = "";

    window.calculationHistory.forEach((item) => {
        const row = document.createElement("div");
        row.className = "history-item";

        const mathDiv = document.createElement("button");
        mathDiv.type = "button";
        mathDiv.className = "history-math-btn";
        mathDiv.title = "Klik for at indlæse udtrykket i feltet igen";
        mathDiv.innerHTML = window.renderLatexToHtml
            ? window.renderLatexToHtml(item.fullLatex, false)
            : item.fullFallbackText;
        mathDiv.addEventListener("click", () => {
            const inputEl = document.getElementById("mathInput");
            if (inputEl) {
                inputEl.value = item.rawInput;
                panelInputDirty = true;
                window.updateMathPreview();
                renderResultInPanel(item, true);
            }
        });

        const insertBtn = document.createElement("button");
        insertBtn.type = "button";
        insertBtn.className = "history-insert-btn secondary-btn";
        insertBtn.title = "Indsæt denne udregning i Word";
        insertBtn.textContent = "Indsæt";
        insertBtn.addEventListener("click", () => {
            window.insertCalcObjectToWord(item);
        });

        row.appendChild(mathDiv);
        row.appendChild(insertBtn);
        container.appendChild(row);
    });
}

/**
 * Henter markeret tekst eller rigtig Word-ligning (<m:oMath>) fra Word ind i panelets felt.
 */
window.fetchFromWordToPanel = async function() {
    const errorBox = document.getElementById("inlineError");
    if (errorBox) errorBox.textContent = "";

    if (!window.isOfficeReady || !window.Word) {
        if (errorBox) errorBox.textContent = "Word-forbindelsen er ikke aktiv (kører i browser-tilstand).";
        return;
    }

    try {
        await Word.run(async (context) => {
            const wordData = await window.WordBridge.readExpressionFromWord(context);
            if (!wordData.expr) {
                if (errorBox) errorBox.textContent = "Ingen ligning eller tekst fundet ved markøren i Word.";
                return;
            }

            const inputEl = document.getElementById("mathInput");
            if (inputEl) {
                inputEl.value = wordData.expr;
                panelInputDirty = true;
                window.updateMathPreview();
                inputEl.focus();
            }
        });
    } catch (err) {
        console.error(err);
        if (errorBox) errorBox.textContent = "Kunne ikke læse fra Word: " + err.message;
    }
};

window.insertCalcObjectToWord = async function(calcObj) {
    const targetCalc = calcObj || window.lastCalculation;
    if (!targetCalc) return;

    const errorBox = document.getElementById("inlineError");
    if (errorBox) errorBox.textContent = "";

    if (!window.isOfficeReady || !window.Word) {
        if (errorBox) errorBox.textContent = "Kan ikke indsætte i Word, da panelet ikke kører inde i Word.";
        return;
    }

    try {
        await Word.run(async (context) => {
            const selection = context.document.getSelection();
            await window.WordBridge.insertMathInWord(
                context,
                selection,
                targetCalc.fullLatex,
                targetCalc.fullFallbackText
            );
            await window.ensureFooter(context);
        });
    } catch (error) {
        console.error(error);
        if (errorBox) errorBox.textContent = "Kunne ikke indsætte i Word: " + error.message;
    }
};

window.calculateInline = async function(actionType) {
    if (isCalculatingInline) return;
    isCalculatingInline = true;

    const errorBox = document.getElementById("inlineError");
    if (errorBox) errorBox.textContent = "";

    const inputEl = document.getElementById("mathInput");
    const panelText = inputEl ? inputEl.value.trim() : "";
    const opts = window.getCalcOptions();
    const autoInsertToggle = document.getElementById("autoInsertToggle");
    const autoInsert = autoInsertToggle ? autoInsertToggle.checked : true;

    try {
        if (!window.isOfficeReady || !window.Word) {
            const calcObj = window.CAS.execute(panelText, actionType, opts);
            renderResultInPanel(calcObj);
            return;
        }

        await Word.run(async (context) => {
            const selection = context.document.getSelection();
            const wordData = await window.WordBridge.readExpressionFromWord(context);

            let sourceMode = "panel";
            let exprToCalculate = panelText;

            if (panelInputDirty && panelText) {
                sourceMode = "panel";
                exprToCalculate = panelText;
            } else if (wordData.sourceMode === "word-highlight" && wordData.expr) {
                sourceMode = "word-highlight";
                exprToCalculate = wordData.expr;
            } else if (panelText) {
                sourceMode = "panel";
                exprToCalculate = panelText;
            } else if (wordData.sourceMode === "word-line" && wordData.expr) {
                sourceMode = "word-line";
                exprToCalculate = wordData.expr;
            } else {
                throw new Error("Skriv et udtryk i feltet eller markér en ligning i Word først.");
            }
            panelInputDirty = false;

            if (inputEl && inputEl.value.trim() !== exprToCalculate) {
                inputEl.value = exprToCalculate;
            }
            window.updateMathPreview();

            const calcObj = window.CAS.execute(exprToCalculate, actionType, opts);
            renderResultInPanel(calcObj);

            if (!autoInsert) {
                return;
            }

            if (sourceMode === "word-highlight") {
                await window.WordBridge.insertMathInWord(
                    context,
                    selection,
                    calcObj.fullLatex,
                    calcObj.fullFallbackText
                );
                await window.ensureFooter(context);
            } else if (sourceMode === "word-line") {
                if (wordData.hasColonPrefix) {
                    // Bevar præfikset før kolon (f.eks. "Opgave 2b: ") og indsæt kun svaret ved markøren
                    await window.WordBridge.insertMathInWord(
                        context,
                        selection,
                        calcObj.answerOnlyLatex,
                        " " + calcObj.fallbackText
                    );
                } else {
                    const lineParagraph = context.document.getSelection().paragraphs.getFirst();
                    const lineRange = lineParagraph.getRange("Content");
                    lineRange.select();
                    await context.sync();

                    await window.WordBridge.insertMathInWord(
                        context,
                        lineRange,
                        calcObj.fullLatex,
                        calcObj.fullFallbackText
                    );
                }
                await window.ensureFooter(context);
            } else if (sourceMode === "panel") {
                await window.WordBridge.insertMathInWord(
                    context,
                    selection,
                    calcObj.fullLatex,
                    calcObj.fullFallbackText
                );
                await window.ensureFooter(context);
            }
        });
    } catch (error) {
        console.error(error);
        if (errorBox) errorBox.textContent = error.message;
    } finally {
        isCalculatingInline = false;
    }
};

document.addEventListener("DOMContentLoaded", function() {
    const inputEl = document.getElementById("mathInput");
    if (inputEl) {
        inputEl.addEventListener("input", function() {
            panelInputDirty = true;
            window.updateMathPreview();
        });
        inputEl.addEventListener("keydown", function(e) {
            if (e.key === "Enter" && !e.shiftKey) {
                e.preventDefault();
                const raw = inputEl.value.trim();
                const isFnOrYDef = /^([a-zA-Z]\s*\(\s*[a-zA-Z]\s*\)|y)\s*=/i.test(raw);
                if (raw.includes("=") && !raw.endsWith("=") && !isFnOrYDef) {
                    window.calculateInline('solve');
                } else {
                    window.calculateInline('eval');
                }
            }
        });
    }
});

