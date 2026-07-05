// ==========================================
// INLINE FUNKTIONER (Dokument-integration) & KLADDEHÆFTE
// ==========================================
window.calculateInline = async function(actionType) {
    // C1: null-check på errorBox så vi ikke crasher hvis elementet mangler
    const errorBox = document.getElementById("inlineError");
    if (errorBox) errorBox.innerText = "";

    try {
        await Word.run(async (context) => {
            const selection = context.document.getSelection();
            selection.load("text");
            await context.sync();

            let textToParse = selection.text.trim();
            let wasHighlighted = true;

            if (!textToParse) {
                wasHighlighted = false;
                const paragraph = selection.paragraphs.getFirst();
                const rangeBeforeCursor = selection.expandTo(paragraph.getRange("Start"));
                rangeBeforeCursor.load("text");
                await context.sync();
                
                textToParse = rangeBeforeCursor.text.trim();
                
                if (textToParse.includes(":")) {
                    textToParse = textToParse.substring(textToParse.lastIndexOf(":") + 1).trim();
                }
            }

            if (!textToParse) {
                throw new Error("Kunne ikke finde noget at regne på. Markér funktionen/ligningen først.");
            }

            // Læs indstillinger fra UI
            const exactToggle = document.getElementById("exactToggle");
            const decimalSelect = document.getElementById("decimalSelect");
            const separatorSelect = document.getElementById("decimalSeparator");
            
            const isExact = exactToggle ? exactToggle.checked : false;
            const decimals = decimalSelect ? parseInt(decimalSelect.value) : 2;
            const decSep = separatorSelect ? separatorSelect.value : ",";

            // Normalisering: brug den delte hjælpefunktion fra utils.js
            // (erstatter π, Π og √ – logikken er samlet ét sted i stedet for at duplikere den)
            let mathExpr = window.normalizeInput(textToParse);
            
            // Konverter komma til punktum for nerdamer (kun hvis de står mellem to tal)
            if (decSep === ",") {
                mathExpr = mathExpr.replace(/(\d),(\d)/g, '$1.$2');
            }
            
            let resultText = "";
            let latexResult = "";
            let originalLatex = "";
            let prefix = "";
            let base64Image = null;

            try {
                // Prøv at lade nerdamer formatere det oprindelige udtryk (laver f.eks. 1/2 til en rigtig brøk)
                try {
                    // Brug nerdamer til at bygge en LaTeX streng, fx "\frac{1}{2}" i stedet for "1/2"
                    originalLatex = nerdamer(mathExpr).toTeX();
                } catch(e) {
                    originalLatex = textToParse; // Fallback til rå tekst
                }

                let mainVar = 'x';
                try {
                    let eqForVars = mathExpr.includes('=') ? mathExpr.split('=')[0] : mathExpr;
                    let vars = nerdamer(eqForVars).variables();
                    if (vars.length > 0) mainVar = vars[0];
                } catch(e) {
                    // Fallback: rens alle kendte funktionsnavne væk først, så finder vi den faktiske variabel.
                    // Den gamle metode (ekskluder enkeltbogstaver) ramte f.eks. 'a' i tan(x)+a.
                    const fnNames = ['sin','cos','tan','log','exp','sqrt','abs','pi','ln'];
                    let cleaned = mathExpr;
                    fnNames.forEach(fn => { cleaned = cleaned.replace(new RegExp(fn, 'g'), ''); });
                    const match = cleaned.match(/[a-zA-Z]/);
                    if (match) mainVar = match[0];
                }

                if (actionType === 'solve') {
                    const parsed = nerdamer.solveEquations(mathExpr, mainVar);
                    let texArray = Array.isArray(parsed) ? parsed.map(val => nerdamer(val).toTeX()) : [nerdamer(parsed).toTeX()];
                    latexResult = texArray.join(" \\lor ");
                    prefix = `\\Rightarrow ${mainVar} = `; 
                } 
                else if (actionType === 'eval' || actionType === 'approx' || actionType === 'calc') {
                    if (isExact) {
                        latexResult = nerdamer(mathExpr).evaluate().toTeX();
                        prefix = "= ";
                    } else {
                        // Afrundet / kommatal
                        let decimalStr = nerdamer(mathExpr).evaluate().text('decimals');
                        latexResult = nerdamer(decimalStr).toTeX('decimals');
                        let originalLatexResult = latexResult;
                        
                        // Afrund til det valgte antal decimaler med safeRound
                        // (undgår IEEE 754 fejl som (1.005).toFixed(2) → "1.00")
                        const regex = new RegExp(`\\d+\\.\\d{${decimals + 1},}`, 'g');
                        latexResult = latexResult.replace(regex, function(match) {
                            const rounded = window.safeRound(parseFloat(match), decimals);
                            return rounded.toFixed(decimals)
                                .replace(/(\.[0-9]*[1-9])0+$/, '$1') // Fjern trailing nuls
                                .replace(/\.0*$/, '');               // Fjern tom decimal
                        });
                        
                        // Tjek om LaTeX-resultatet indeholder 10+ decimaler (f.eks. pi eller 1/3)
                        // RETTET: den gamle regex /\\.\\d{10,}/ matchede bogstaveligt \\ og \\d,
                        // ikke et punktum efterfulgt af cifre. Korrekt regex er /\.\d{10,}/
                        const hasManyDecimals = /\.\d{10,}/.test(originalLatexResult);
                        if (originalLatexResult !== latexResult || hasManyDecimals) {
                            prefix = "\\approx ";
                        } else {
                            prefix = "= ";
                        }
                    }
                }
                else if (actionType === 'diff') {
                    latexResult = nerdamer('diff(' + mathExpr + ', ' + mainVar + ')').evaluate().toTeX();
                    prefix = `\\Rightarrow f'(${mainVar}) = `;
                }
                else if (actionType === 'int') {
                    latexResult = nerdamer('integrate(' + mathExpr + ', ' + mainVar + ')').toTeX();
                    prefix = `\\Rightarrow \\int f(${mainVar}) d${mainVar} = `;
                }
                // Graf-handling er fjernet: 'graph'-action-type eksisterede aldrig
                // i UI'et og kaldte den udefinerede drawGraphAndGetBase64().
                // Grafer indsættes via GeoGebra-modalen i graphing.js.
                console.warn("actionType 'graph' er ikke understøttet fra calculateInline.");
            } catch (calcError) {
                // Videresend den konkrete fejlbesked fra nerdamer så brugeren får en meningsfyldt fejl.
                // Den generiske besked skjulte f.eks. "division by zero" og "unsupported operation".
                throw new Error(`Matematisk fejl: ${calcError.message}. Tjek at udtrykket er gyldigt og at variablen hedder 'x' (ikke t, k etc.).`);
            }
            
            // Udskift punktum med komma i outputtet, hvis indstillingen er sat til komma
            if (decSep === "," && latexResult) {
                latexResult = latexResult.replace(/(\d)\.(\d)/g, '$1,$2');
            }

            // Generel formatering af resultatet
            if (latexResult) {
                // Fjern \cdot foran \pi for et mere naturligt udseende (f.eks. 2\pi i stedet for 2 \cdot \pi)
                latexResult = latexResult.replace(/\\cdot\s*\\pi/g, '\\pi');
            }

            if (actionType === 'graph' && base64Image) {
                let p = selection.insertParagraph("", "After");
                p.insertInlinePictureFromBase64(base64Image, "End");
                await context.sync();
            } else {
                // Hjælpefunktion hvis vi falder tilbage til at indsætte rå tekst
                const getFallbackText = () => {
                    let text = prefix + latexResult;
                    return text
                        .replace(/\\approx\s*/g, '≈ ')
                        .replace(/\\Rightarrow\s*/g, '⇒ ')
                        .replace(/\\int\s*/g, '∫ ')
                        .replace(/\\pi/g, 'π')
                        .replace(/\\cdot/g, '·')
                        .replace(/\\lor/g, ' v ')
                        .replace(/\\/g, ''); // Fjern alle andre backslashes for et renere look
                };

                let ooxml = "";
                try {
                    let targetRange = context.document.getSelection();
                    
                    if (window.latexToOmmlString) {
                        const stackedToggle = document.getElementById("stackedFractionToggle");
                        const isStacked = stackedToggle ? stackedToggle.checked : true;
                        
                        let fullLatex = wasHighlighted 
                            ? originalLatex + " " + prefix + latexResult
                            : prefix + latexResult;
                            
                        if (window.formatFractions) {
                            fullLatex = window.formatFractions(fullLatex, isStacked);
                        }
                            
                        let ommlString = window.latexToOmmlString(fullLatex);
                        
                        // Hvis brugeren har markeret tekst, prøver vi at indsnævre markeringen for at undgå at slette usynlige linjeskift (som ødelægger linjen ovenover)
                        if (wasHighlighted && textToParse) {
                            let searchResults = targetRange.search(textToParse, {matchCase: true});
                            searchResults.load("items");
                            await context.sync();
                            if (searchResults.items.length > 0) {
                                targetRange = searchResults.items[0];
                            }
                        }
                        
                        // 1. Indpak markeringen i en midlertidig Content Control
                        let ccTag = "MATH_TARGET_" + Date.now();
                        let cc = targetRange.insertContentControl();
                        cc.tag = ccTag;
                        
                        // Sørg for at CC ikke er tom, ellers kan Word Web fjerne den fra OOXML. 
                        // Det overskriver kun originalteksten, som vi alligevel skal erstatte.
                        cc.insertText("MATH_PLACEHOLDER", "Replace");
                        
                        // 2. Find afsnittet
                        let paragraph = cc.paragraphs.getFirst();
                        let pOoxml = paragraph.getOoxml();
                        await context.sync();
                        
                        // 3. Parse afsnittets OOXML
                        let parser = new DOMParser();
                        let xmlDoc = parser.parseFromString(pOoxml.value, "text/xml");
                        
                        // 4. Find vores Content Control tag (<w:sdt>)
                        let targetSdt = null;
                        let tags = xmlDoc.getElementsByTagName("*");
                        
                        // A: Prøv via ccTag
                        for (let i = 0; i < tags.length; i++) {
                            if (tags[i].nodeName === "w:tag" || tags[i].localName === "tag") {
                                let val = tags[i].getAttribute("w:val") || tags[i].getAttribute("val");
                                if (val === ccTag) {
                                    targetSdt = tags[i].parentNode.parentNode;
                                    break;
                                }
                            }
                        }
                        
                        // B: Fallback via MATH_PLACEHOLDER (hvis tag blev strippet)
                        if (!targetSdt) {
                            for (let i = 0; i < tags.length; i++) {
                                if ((tags[i].nodeName === "w:t" || tags[i].localName === "t") && tags[i].textContent.includes("MATH_PLACEHOLDER")) {
                                    let p = tags[i].parentNode;
                                    while (p && p.nodeName !== "w:sdt" && p.localName !== "sdt") {
                                        p = p.parentNode;
                                    }
                                    if (p) {
                                        targetSdt = p;
                                        break;
                                    }
                                }
                            }
                        }
                        
                        // C: Sidste udvej (find en sdt)
                        if (!targetSdt) {
                            let sdts = xmlDoc.getElementsByTagName("w:sdt");
                            if (sdts.length === 0) sdts = xmlDoc.getElementsByTagName("sdt");
                            if (sdts.length > 0) targetSdt = sdts[sdts.length - 1];
                        }
                        
                        if (targetSdt) {
                            // 5. Parse OMML
                            let mathDoc = parser.parseFromString(ommlString, "text/xml");
                            let importedMathNode = xmlDoc.importNode(mathDoc.documentElement, true);
                            
                            // 6. Erstat Content Control (<w:sdt>) med vores <m:oMath>
                            targetSdt.parentNode.replaceChild(importedMathNode, targetSdt);
                            
                            // 7. Serialisér tilbage til OOXML string
                            let serializer = new XMLSerializer();
                            let newOoxml = serializer.serializeToString(xmlDoc);
                            
                            // 8. Erstat hele afsnittet med den modifikerede OOXML (blok-til-blok erstatning virker i Word Web)
                            paragraph.insertOoxml(newOoxml, "Replace");
                            await context.sync(); 
                        } else {
                            // Fallback hvis vi ikke kunne finde content controlen (sker sjældent)
                            targetRange.insertOoxml(window.latexToOoxml(fullLatex), "Replace");
                            await context.sync();
                        }
                    } else {
                        const fallbackTxt = wasHighlighted ? textToParse + " " + getFallbackText() : " " + getFallbackText();
                        
                        if (wasHighlighted && textToParse) {
                            let searchResults = targetRange.search(textToParse, {matchCase: true});
                            searchResults.load("items");
                            await context.sync();
                            if (searchResults.items.length > 0) {
                                targetRange = searchResults.items[0];
                            }
                        }
                        
                        targetRange.insertText(fallbackTxt, "Replace");
                        await context.sync();
                    }
                } catch (ooxmlErr) {
                    console.error("❌ OOXML-indsættelse fejlede:", ooxmlErr);
                    console.error("   Fejlbesked:", ooxmlErr.message);
                    console.log("↩️ Bruger tekst-fallback efter OOXML-fejl");
                    
                    let insertRange = context.document.getSelection();
                    const fallbackTxt = wasHighlighted ? textToParse + " " + getFallbackText() : " " + getFallbackText();
                    
                    if (wasHighlighted && textToParse) {
                        let searchResults = insertRange.search(textToParse, {matchCase: true});
                        searchResults.load("items");
                        await context.sync();
                        if (searchResults.items.length > 0) {
                            insertRange = searchResults.items[0];
                        }
                    }
                    
                    insertRange.insertText(fallbackTxt, "Replace");
                    await context.sync();
                }
            }
            
            await window.ensureFooter(context);
        });
    } catch (error) {
        console.error(error);
        errorBox.innerText = error.message;
    }
}

window.calculateDraft = function() {
    const expression = document.getElementById("expression").value.trim();
    const resultContainer = document.getElementById("resultContainer");
    const insertBtn = document.getElementById("insertBtn");

    if (!expression) {
        if (resultContainer) resultContainer.innerText = "Skriv et udtryk først.";
        return;
    }

    const exactToggle = document.getElementById("exactToggle");
    const decimalSelect = document.getElementById("decimalSelect");
    const separatorSelect = document.getElementById("decimalSeparator");
    
    const isExact = exactToggle ? exactToggle.checked : false;
    const decimals = decimalSelect ? parseInt(decimalSelect.value) : 2;
    const decSep = separatorSelect ? separatorSelect.value : ",";

    try {
        let mathExpr = window.normalizeInput(expression);
        let result = "";
        if (mathExpr.includes("=")) {
            const parsed = nerdamer.solveEquations(mathExpr);
            result = Array.isArray(parsed) ? parsed.map(val => val.toString()).join(", ") : parsed.toString();
        } else {
            // Konverter input til punktum for nerdamer
            if (decSep === ",") {
                mathExpr = mathExpr.replace(/(\d),(\d)/g, '$1.$2');
            }

            if (isExact) {
                result = nerdamer(mathExpr).evaluate().text();
            } else {
                result = nerdamer(mathExpr).evaluate().text('decimals');
                const regex = new RegExp(`\\d+\\.\\d{${decimals + 1},}`, 'g');
                result = result.replace(regex, function(match) {
                    const rounded = window.safeRound(parseFloat(match), decimals);
                    return rounded.toFixed(decimals)
                        .replace(/(\.[0-9]*[1-9])0+$/, '$1')
                        .replace(/\.0*$/, '');
                });
            }
        }
        
        // Formater resultatet (fjern f.eks. * foran pi og variabler for pænere output i kladdehæftet)
        if (typeof result === 'string') {
            result = result.replace(/\*pi/g, 'π');
            result = result.replace(/\*([a-zA-Z])/g, '$1'); 
            result = result.replace(/pi/g, 'π');
        }
        
        if (decSep === "," && typeof result === 'string') {
            result = result.replace(/(\d)\.(\d)/g, '$1,$2');
        }

        window.currentResult = " = " + result;
        resultContainer.innerText = window.currentResult;
        insertBtn.disabled = false;
    } catch (error) {
        resultContainer.innerText = "Fejl: Kunne ikke beregne.";
        insertBtn.disabled = true;
    }
}
