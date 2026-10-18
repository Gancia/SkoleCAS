// ============================================================================
// SkoleCAS Kerne:
//  1. window.CAS           – Udskiftelig CAS-adapter (eval, solve, diff, int)
//  2. window.WordBridge    – To-vejs bro til Word (læser OG skriver <m:oMath> OMML)
//  3. window.GeoGebraBridge– Lazy-loaded GeoGebra-graftegner & ekstern bro
// ============================================================================

window.calculationHistory = [];
window.lastCalculation = null;
let panelInputDirty = false;

// ─────────────────────────────────────────────────────────────────
//  1. CAS-ADAPTER (Isoleret beregningslag)
// ─────────────────────────────────────────────────────────────────

function detectMainVariable(mathExpr) {
    try {
        const exprForVars = mathExpr.includes('=') ? mathExpr.split('=')[0] : mathExpr;
        const vars = nerdamer(exprForVars).variables();
        if (Array.isArray(vars) && vars.length > 0) {
            if (vars.includes('x')) return 'x';
            return vars[0];
        }
    } catch (e) {
        // Fallback nedenfor
    }

    const fnNames = [
        'arcsin', 'arccos', 'arctan', 'sinh', 'cosh', 'tanh',
        'asin', 'acos', 'atan', 'sqrt', 'sin', 'cos', 'tan',
        'cot', 'sec', 'csc', 'log', 'exp', 'abs', 'pi', 'ln'
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
    const roundedStr = str.replace(/-?\d+\.\d+/g, function(match) {
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
     * Konverterer et brugerudtryk til LaTeX til forhåndsvisning.
     */
    toPreviewLatex: function(rawInput, decSep, isStacked) {
        if (!rawInput || !rawInput.trim()) return "";
        let cleaned = rawInput.trim();
        if (cleaned.includes('⇒')) cleaned = cleaned.split('⇒')[0].trim();

        const sep = decSep || ",";
        const stacked = isStacked !== undefined ? isStacked : true;

        try {
            let norm = window.normalizeInput(cleaned);
            norm = window.convertDanishDecimalToDot(norm, sep);

            if (norm.includes('=')) {
                const parts = norm.split('=');
                const leftRaw = parts[0].trim();
                const rightRaw = parts.slice(1).join('=').trim();

                let leftTex = leftRaw;
                if (!/^[a-zA-Z]\s*\([a-zA-Z]\)$/.test(leftRaw)) {
                    try { leftTex = nerdamer(leftRaw).toTeX(); } catch (e) { leftTex = leftRaw; }
                }
                let rightTex = rightRaw ? nerdamer(rightRaw).toTeX() : "";
                let combined = rightRaw ? `${leftTex} = ${rightTex}` : `${leftTex} =`;
                if (window.formatFractions) combined = window.formatFractions(combined, stacked);
                if (sep === ",") combined = combined.replace(/(\d)\.(\d)/g, '$1,$2');
                return combined.replace(/\\cdot\s*\\pi/g, '\\pi');
            }

            let tex = nerdamer(norm).toTeX();
            if (window.formatFractions) tex = window.formatFractions(tex, stacked);
            if (sep === ",") tex = tex.replace(/(\d)\.(\d)/g, '$1,$2');
            return tex.replace(/\\cdot\s*\\pi/g, '\\pi');
        } catch (e) {
            let fallback = cleaned
                .replace(/[πΠ]/g, '\\pi ')
                .replace(/√\(([^)]*)\)/g, '\\sqrt{$1}')
                .replace(/\*/g, ' \\cdot ');
            if (window.formatFractions) fallback = window.formatFractions(fallback, stacked);
            return fallback;
        }
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
                const exprToDiff = fnHeaderMatch ? fnHeaderMatch[3] : (yHeaderMatch ? yHeaderMatch[1] : mathExpr);
                const fnName = fnHeaderMatch ? fnHeaderMatch[1] : 'f';
                latexResult = nerdamer(`diff(${exprToDiff}, ${mainVar})`).evaluate().toTeX();
                prefix = `\\Rightarrow ${fnName}'(${mainVar}) = `;
            }
            else if (actionType === 'int') {
                const exprToInt = fnHeaderMatch ? fnHeaderMatch[3] : (yHeaderMatch ? yHeaderMatch[1] : mathExpr);
                latexResult = nerdamer(`integrate(${exprToInt}, ${mainVar})`).toTeX() + " + C";
                prefix = `\\Rightarrow \\int f(${mainVar})\\,d${mainVar} = `;
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

// Bagudkompatible globale funktioner
window.expressionToPreviewLatex = window.CAS.toPreviewLatex;
window.evaluateMath = window.CAS.execute;

// ─────────────────────────────────────────────────────────────────
//  2. WORDBRIDGE (To-vejs læsning & skrivning af OMML i Word)
// ─────────────────────────────────────────────────────────────────

window.WordBridge = {
    /**
     * Læser matematisk udtryk fra Word (både rigtige <m:oMath>-ligninger og rå tekst!).
     * Returnerer { expr: string, sourceMode: 'word-highlight'|'word-line'|null, rawHighlightedText: string }
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
                    rawHighlightedText: selection.text ? selection.text.trim() : ""
                };
            }
        }

        // 2. Tjek om brugeren har markeret almindelig tekst
        const highlightedText = selection.text ? selection.text.trim() : "";
        if (highlightedText) {
            return {
                expr: highlightedText,
                sourceMode: "word-highlight",
                rawHighlightedText: highlightedText
            };
        }

        // 3. Ingen markering: Tjek aktuelt afsnit for enten en <m:oMath>-ligning eller tekst før markøren
        const paragraph = selection.paragraphs.getFirst();
        const rangeBeforeCursor = selection.expandTo(paragraph.getRange("Start"));
        rangeBeforeCursor.load("text");
        const paraOoxml = paragraph.getOoxml();
        await context.sync();

        if (window.ommlToExpression && paraOoxml && paraOoxml.value) {
            const paraOmmlExpr = window.ommlToExpression(paraOoxml.value);
            if (paraOmmlExpr) {
                return {
                    expr: paraOmmlExpr,
                    sourceMode: "word-line",
                    rawHighlightedText: ""
                };
            }
        }

        let lineText = rangeBeforeCursor.text ? rangeBeforeCursor.text.trim() : "";
        if (lineText.includes(":")) {
            lineText = lineText.substring(lineText.lastIndexOf(":") + 1).trim();
        }
        if (lineText) {
            return {
                expr: lineText,
                sourceMode: "word-line",
                rawHighlightedText: ""
            };
        }

        return { expr: "", sourceMode: null, rawHighlightedText: "" };
    },

    /**
     * Indsætter et matematisk resultat som rigtig Word-ligning (<m:oMath>) uden låste bokse.
     */
    insertMathInWord: async function(context, targetRange, latexToInsert, fallbackTextToInsert, wasHighlighted, originalTextToMatch) {
        if (wasHighlighted && originalTextToMatch) {
            try {
                const searchResults = targetRange.search(originalTextToMatch, { matchCase: true });
                searchResults.load("items");
                await context.sync();
                if (searchResults.items.length > 0) {
                    targetRange = searchResults.items[0];
                }
            } catch (e) {
                // Ignorér hvis søgetekst indeholder specialtegn
            }
        }

        if (window.latexToOmmlString) {
            try {
                const ommlString = window.latexToOmmlString(latexToInsert);
                const ccTag = "MATH_TARGET_" + Date.now();
                const cc = targetRange.insertContentControl();
                cc.tag = ccTag;
                cc.cannotDelete = false;
                cc.cannotEdit = false;
                cc.insertText("[Beregner...]", "Replace");

                const paragraph = cc.paragraphs.getFirst();
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

                let replacedInXml = false;
                if (sdtNode) {
                    let sdtContent = null;
                    for (let i = 0; i < sdtNode.childNodes.length; i++) {
                        const child = sdtNode.childNodes[i];
                        if (child.nodeName === "w:sdtContent" || child.localName === "sdtContent") {
                            sdtContent = child;
                            break;
                        }
                    }
                    if (sdtContent) {
                        while (sdtContent.firstChild) {
                            sdtContent.removeChild(sdtContent.firstChild);
                        }
                        sdtContent.appendChild(importedMathNode);
                        replacedInXml = true;
                    } else {
                        sdtNode.parentNode.replaceChild(importedMathNode, sdtNode);
                        replacedInXml = true;
                    }
                } else if (rNode) {
                    rNode.parentNode.replaceChild(importedMathNode, rNode);
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

                    // Pak Content Control-rammen rent ud med delete(true), så ligningen aldrig låser i Word Web
                    const remainingCcs = paragraph.contentControls;
                    remainingCcs.load("items/tag");
                    await context.sync();

                    for (let i = 0; i < remainingCcs.items.length; i++) {
                        const item = remainingCcs.items[i];
                        if (!item.tag || item.tag.startsWith("MATH_TARGET_")) {
                            item.cannotDelete = false;
                            item.cannotEdit = false;
                            item.delete(true);
                        }
                    }
                    await context.sync();
                    return;
                }
            } catch (ooxmlErr) {
                console.warn("OOXML-indsættelse fejlede, bruger tekst-fallback:", ooxmlErr);
            }
        }

        targetRange.insertText(fallbackTextToInsert, "Replace");
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

function renderResultInPanel(calcObj) {
    window.lastCalculation = calcObj;
    window.currentResult = " " + calcObj.fallbackText;

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
                renderResultInPanel(item);
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
            selection.load("text");
            await context.sync();

            const hasSelection = selection.text.trim().length > 0;
            await window.WordBridge.insertMathInWord(
                context,
                selection,
                targetCalc.fullLatex,
                targetCalc.fullFallbackText,
                hasSelection,
                selection.text.trim()
            );
            await window.ensureFooter(context);
        });
    } catch (error) {
        console.error(error);
        if (errorBox) errorBox.textContent = "Kunne ikke indsætte i Word: " + error.message;
    }
};

window.calculateInline = async function(actionType) {
    const errorBox = document.getElementById("inlineError");
    if (errorBox) errorBox.textContent = "";

    const inputEl = document.getElementById("mathInput");
    const panelText = inputEl ? inputEl.value.trim() : "";
    const opts = window.getCalcOptions();
    const autoInsertToggle = document.getElementById("autoInsertToggle");
    const autoInsert = autoInsertToggle ? autoInsertToggle.checked : true;

    if (!window.isOfficeReady || !window.Word) {
        try {
            const calcObj = window.CAS.execute(panelText, actionType, opts);
            renderResultInPanel(calcObj);
        } catch (err) {
            if (errorBox) errorBox.textContent = err.message;
        }
        return;
    }

    try {
        await Word.run(async (context) => {
            const selection = context.document.getSelection();
            const wordData = await window.WordBridge.readExpressionFromWord(context);

            let sourceMode = "panel";
            let exprToCalculate = panelText;

            if (wordData.sourceMode === "word-highlight" && wordData.expr) {
                sourceMode = "word-highlight";
                exprToCalculate = wordData.expr;
            } else if (panelInputDirty && panelText) {
                sourceMode = "panel";
                exprToCalculate = panelText;
            } else if (wordData.sourceMode === "word-line" && wordData.expr) {
                sourceMode = "word-line";
                exprToCalculate = wordData.expr;
            } else if (panelText) {
                sourceMode = "panel";
                exprToCalculate = panelText;
            } else {
                throw new Error("Skriv et udtryk i feltet eller markér en ligning i Word først.");
            }

            if (inputEl && inputEl.value.trim() !== exprToCalculate) {
                inputEl.value = exprToCalculate;
                window.updateMathPreview();
            }
            panelInputDirty = false;

            const calcObj = window.CAS.execute(exprToCalculate, actionType, opts);
            renderResultInPanel(calcObj);

            if (sourceMode === "word-highlight") {
                await window.WordBridge.insertMathInWord(
                    context,
                    selection,
                    calcObj.fullLatex,
                    calcObj.fullFallbackText,
                    Boolean(wordData.rawHighlightedText),
                    wordData.rawHighlightedText
                );
                await window.ensureFooter(context);
            } else if (sourceMode === "word-line") {
                await window.WordBridge.insertMathInWord(
                    context,
                    selection,
                    calcObj.answerOnlyLatex,
                    " " + calcObj.fallbackText,
                    false,
                    ""
                );
                await window.ensureFooter(context);
            } else if (sourceMode === "panel" && autoInsert) {
                await window.WordBridge.insertMathInWord(
                    context,
                    selection,
                    calcObj.fullLatex,
                    calcObj.fullFallbackText,
                    false,
                    ""
                );
                await window.ensureFooter(context);
            }
        });
    } catch (error) {
        console.error(error);
        if (errorBox) errorBox.textContent = error.message;
    }
};

window.calculateDraft = function() {
    window.calculateInline('eval');
};
window.insertToWord = function() {
    window.insertCalcObjectToWord(window.lastCalculation);
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
                if (raw.includes("=")) {
                    window.calculateInline('solve');
                } else {
                    window.calculateInline('eval');
                }
            }
        });
    }
});
