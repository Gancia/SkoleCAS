// ==========================================
// REGRESSION (GeoGebra Integration)
// ==========================================

window.calculateRegression = async function() {
    const errorBox = document.getElementById("regError");
    if (errorBox) errorBox.innerText = "";

    // A3: Null-checks på DOM-elementer så vi ikke crasher ved HTML-ændringer
    const inputEl = document.getElementById("regInput");
    const typeEl  = document.getElementById("regType");
    if (!inputEl || !typeEl) {
        if (errorBox) errorBox.innerText = "Intern fejl: kan ikke finde inputfelterne.";
        return;
    }
    const input = inputEl.value.trim();
    const type  = typeEl.value;
    
    if (!input) {
        if (errorBox) errorBox.innerText = "Indsæt venligst data først.";
        return;
    }

    // Parse data
    const lines = input.split('\n');
    let points = [];
    for (let line of lines) {
        line = line.trim();
        if (!line) continue;
        
        // Håndter tab, semikolon, flere mellemrum samt evt. copy-paste snavs
        let parts = line.split(/[\s\t;|]+/);
        let nums = [];
        
        for (let p of parts) {
            let clean = p.replace(/['"]/g, '').replace(/,/g, '.');
            let val = parseFloat(clean);
            if (!isNaN(val)) {
                nums.push(val);
            }
        }
        
        if (nums.length >= 2) {
            points.push(`(${nums[0]}, ${nums[1]})`);
        }
    }

    if (points.length < 2) {
        if (errorBox) errorBox.innerText = "Der skal mindst være to gyldige punkter (x og y).";
        return;
    }

    const listStr = `L_1 = {${points.join(', ')}}`;

    if (!window.ggbApplet) {
        // Åbn GeoGebra-modalen i baggrunden og bed brugeren vente
        if (window.openGgbModal) window.openGgbModal();
        if (errorBox) errorBox.innerText = "GeoGebra starter... Vent ca. 5 sekunder og prøv igen.";
        return;
    }

    try {
        // Nulstil
        window.ggbApplet.evalCommand("Delete(L_1)");
        window.ggbApplet.evalCommand("Delete(f)");
        window.ggbApplet.evalCommand("Delete(R)");
        
        // Indsæt liste
        window.ggbApplet.evalCommand(listStr);
        window.ggbApplet.evalCommand("SetVisibleInView(L_1, 1, true)");
        window.ggbApplet.evalCommand("ZoomIn(L_1)");

        // Beregn regression
        let fitCmd = "FitLine";
        if (type === "Eksponentiel") fitCmd = "FitGrowth"; // Bedre end FitExp til b * a^x (dansk standard)
        else if (type === "Potens") fitCmd = "FitPow";
        else if (type === "Polynomium") fitCmd = "FitPoly(L_1, 2)";

        if (type === "Polynomium") {
            window.ggbApplet.evalCommand(`f = ${fitCmd}`);
        } else {
            window.ggbApplet.evalCommand(`f = ${fitCmd}(L_1)`);
        }
        
        window.ggbApplet.evalCommand("R = RSquare(L_1, f)");

        // Hent LaTeX udtryk
        let fLatexStr = window.ggbApplet.getLaTeXString("f");
        if (!fLatexStr) {
            throw new Error("Kunne ikke beregne funktion. Tjek data og regressionstype.");
        }
        
        // GeoGebra RSquare value
        let rValue = window.ggbApplet.getValue("R");
        // Afrund R^2 til 4 decimaler for pænhed
        let rRounded = isNaN(rValue) ? "N/A" : Number(rValue).toFixed(4).replace('.', ',');
        
        let rLatexStr = `R^2 = ${rRounded}`;
        
        // O2: Erstat KUN cifre-omgivne punktummer (dansk decimal-notation).
        // Den gamle /\.g/ erstattede ALLE punktummer inkl. LaTeX-kommandoer.
        fLatexStr = fLatexStr.replace(/(\d)\.(\d)/g, '$1,$2');
        
        let fullLatex = `f(x) = ${fLatexStr} \\quad \\lor \\quad ${rLatexStr}`;

        // Konverter til Word OOXML
        let ommlString = "";
        let ooxml = "";
        
        const stackedToggle = document.getElementById("stackedFractionToggle");
        const isStacked = stackedToggle ? stackedToggle.checked : true;
        if (window.formatFractions) {
            fullLatex = window.formatFractions(fullLatex, isStacked);
        }
        
        if (window.latexToOmmlString) {
            ommlString = window.latexToOmmlString(fullLatex);
        } else if (window.latexToOoxml) {
            ooxml = window.latexToOoxml(fullLatex); // Fallback
        }

        // O1: Valider at base64-strengen ikke er tom inden indsættelse.
        // GeoGebra returnerer en tom streng hvis der ikke er tegnet noget.
        const base64 = await new Promise((resolve, reject) => {
            try {
                window.ggbApplet.getPNGBase64(1, false, 300, function(data) {
                    if (!data || data.length < 100) {
                        reject(new Error("GeoGebra returnerede et tomt grafbillede — prøv at zoome ind og beregn igen."));
                    } else {
                        resolve(data);
                    }
                });
            } catch (e) {
                reject(e);
            }
        });

        await Word.run(async (context) => {
            const range = context.document.getSelection();
            let insertedParagraph = null;
            
            if (ommlString) {
                // 1. Indpak markeringen i en midlertidig Content Control
                let ccTag = "MATH_TARGET_" + Date.now();
                let cc = range.insertContentControl();
                cc.tag = ccTag;
                
                // Sørg for at CC ikke er tom
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
                
                // B: Fallback via MATH_PLACEHOLDER
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
                
                // C: Sidste udvej
                if (!targetSdt) {
                    let sdts = xmlDoc.getElementsByTagName("w:sdt");
                    if (sdts.length === 0) sdts = xmlDoc.getElementsByTagName("sdt");
                    if (sdts.length > 0) targetSdt = sdts[sdts.length - 1];
                }
                
                if (targetSdt) {
                    let mathDoc = parser.parseFromString(ommlString, "text/xml");
                    let importedMathNode = xmlDoc.importNode(mathDoc.documentElement, true);
                    targetSdt.parentNode.replaceChild(importedMathNode, targetSdt);
                    
                    let serializer = new XMLSerializer();
                    let newOoxml = serializer.serializeToString(xmlDoc);
                    
                    let preambleMatch = pOoxml.value.match(/^(<\\?xml[^>]+>\\s*<\\?mso-application[^>]+>\\s*)/i);
                    let preamble = preambleMatch ? preambleMatch[1] : '<?xml version="1.0" standalone="yes"?>\n<?mso-application progid="Word.Document"?>\n';
                    newOoxml = newOoxml.replace(/^<\\?xml[^>]+>\\s*/i, '');
                    newOoxml = preamble + newOoxml;
                    
                    insertedParagraph = paragraph.insertOoxml(newOoxml, "Replace");
                } else {
                    let insertedRange = range.insertOoxml(window.latexToOoxml(fullLatex), "Replace");
                    insertedParagraph = insertedRange.paragraphs.getLast();
                }
            } else if (ooxml) {
                let insertedRange = range.insertOoxml(ooxml, "Replace");
                insertedParagraph = insertedRange.paragraphs.getLast();
            } else {
                // Fallback
                let fallbackTxt = fullLatex.replace(/\\quad \\lor \\quad/g, ' eller ').replace(/\\cdot/g, '*');
                let insertedRange = range.insertText(fallbackTxt, "Replace");
                insertedParagraph = insertedRange.paragraphs.getLast();
            }
            
            // Indsæt billede efter det netop indsatte/opdaterede afsnit
            let p = insertedParagraph.insertParagraph("", "After");
            p.insertInlinePictureFromBase64(base64, "End");
            
            if (window.ensureFooter) {
                await window.ensureFooter(context);
            }
            await context.sync();
        });

    } catch (err) {
        console.error(err);
        // C5: Vis den faktiske fejlbesked frem for en generisk streng.
        // Den generiske besked skjulte alle fejltyper for brugeren.
        if (errorBox) errorBox.innerText = `Fejl: ${err.message || 'Ukendt fejl — tjek data og regressionstype.'}`;
    }
};
