// ==========================================
// GRAFTEGNER (GeoGebra Modal Integration)
// ==========================================

window.ggbApplet      = null;
window.ggbInitialized  = false;
window.ggbInitializing = false;

/**
 * Initialiserer GeoGebra i #ggb-element (inde i modalen).
 * Kalles kun én gang — ignorerer genkaldt.
 * Bruges også af regression.js som tjekker window.ggbApplet.
 */
window.initGeoGebra = function() {
    if (window.ggbInitialized || window.ggbInitializing) return;
    window.ggbInitializing = true;

    // C4: Hvis GeoGebra-CDN'et fejler (ingen netadgang) bliver ggbInitializing
    // hængende på true for evigt og brugeren kan aldrig retry.
    // Timer nulstiller flaget efter 15s og viser en forståelig fejlbesked.
    const initTimeout = setTimeout(function() {
        if (!window.ggbInitialized) {
            window.ggbInitializing = false;
            setModalStatus('\u274c GeoGebra kunne ikke starte \u2014 tjek internetforbindelsen og pr\u00f8v igen');
        }
    }, 15000);

    // Hent modalens faktiske dimensioner for korrekt GeoGebra-størrelse
    const modalBody = document.querySelector('.ggb-modal-body');
    const w = modalBody ? Math.max(modalBody.clientWidth,  300) : 600;
    const h = modalBody ? Math.max(modalBody.clientHeight, 300) : 450;

    var parameters = {
        "id":               "ggbApplet",
        "width":            w,
        "height":           h,
        "showMenuBar":      false,
        "showAlgebraInput": true,
        "showToolBar":      true,
        "customToolBar":    "0 73 62 | 1 501 67 , 5 19 , 72 75 76 | 2 15 45 , 18 65 , 7 37 | 4 3 8 9 , 13 44 , 58 , 47 | 16 51 64 , 70 | 10 34 53 11 , 24 20 22 , 21 23 | 55 56 57 , 12 | 36 46 , 38 49 50 , 71 14 68 | 30 29 54 32 31 33 | 25 17 26 60 52 61 | 40 41 42 , 27 28 35 , 6",
        "showToolBarHelp":  false,
        "showResetIcon":    true,
        "enableLabelDrags": false,
        "enableShiftDragZoom": true,
        "enableRightClick": true,
        "errorDialogsActive": true,
        "useBrowserForJS":  false,
        "allowStyleBar":    true,
        "preventFocus":     false,
        "showZoomButtons":  true,
        "capturingThreshold": null,
        "appletOnLoad": function(api) {
            clearTimeout(initTimeout);     // Annuller timeout-vagten
            window.ggbApplet      = api;
            window.ggbInitialized  = true;
            window.ggbInitializing = false;
            setModalStatus('GeoGebra klar \u2014 skriv en funktion i algebraboksen');
        },
        "appName": "graphing",
        "language": "da"
    };

    var applet = new GGBApplet(parameters, true);
    applet.inject('ggb-element');
};

// ==========================================
// HJÆLPEFUNKTIONER
// ==========================================

function setModalStatus(msg) {
    const el = document.getElementById('ggbModalStatus');
    if (el) el.textContent = msg;
}

function openGgbModal() {
    const modal = document.getElementById('ggbModal');
    if (!modal) return;
    modal.classList.add('open');

    if (!window.ggbInitialized && !window.ggbInitializing) {
        setModalStatus("Starter GeoGebra...");
        // To frames: sikrer at modalen er synlig (display:flex) inden GeoGebra
        // forsøger at måle #ggb-element-containerens dimensioner.
        requestAnimationFrame(() => requestAnimationFrame(window.initGeoGebra));
    }
}

function closeGgbModal() {
    const modal = document.getElementById('ggbModal');
    if (modal) modal.classList.remove('open');
}

/**
 * Henter markeret tekst fra Word og sender den til GeoGebra.
 * Deles af knappen i Graf-tab OG knappen i modalen.
 */
async function loadEquationFromWord() {
    if (!window.ggbApplet) {
        setModalStatus("GeoGebra er ikke klar endnu \u2014 vent et \u00f8jeblik");
        return;
    }
    try {
        await Word.run(async (context) => {
            const selection = context.document.getSelection();
            selection.load("text");
            await context.sync();

            let text = selection.text.trim();

            if (!text) {
                const paragraph = selection.paragraphs.getFirst();
                const rangeBeforeCursor = selection.expandTo(paragraph.getRange("Start"));
                rangeBeforeCursor.load("text");
                await context.sync();

                text = rangeBeforeCursor.text.trim();
                if (text.includes(":")) {
                    text = text.substring(text.lastIndexOf(":") + 1).trim();
                }
            }

            if (!text) {
                setModalStatus("Ingen tekst markeret i Word");
                return;
            }

            // Normalisér: π → pi, dansk komma → punktum
            let mathExpr = window.normalizeInput(text);
            mathExpr = mathExpr.replace(/(\d),(\d)/g, '$1.$2');

            window.ggbApplet.evalCommand(mathExpr);
            setModalStatus(`Sendt til GeoGebra: ${text}`);
        });
    } catch (e) {
        console.error("Fejl ved hentning af ligning:", e);
        setModalStatus("Fejl ved hentning af ligning fra Word");
    }
}

// ==========================================
// INDSÆT GRAF I WORD
// Primær fix: getPNGBase64 pakkes i en Promise
// og base64-strengen valideres inden indsættelse.
// ==========================================
async function insertGraphToWord() {
    if (!window.ggbApplet) {
        setModalStatus("GeoGebra er ikke klar endnu");
        return;
    }

    setModalStatus("Henter grafbillede...");

    try {
        // Pak GeoGebras callback-baserede API ind i en Promise
        // så async/await og try/catch virker korrekt
        const base64 = await new Promise((resolve, reject) => {
            try {
                window.ggbApplet.getPNGBase64(1, false, 150, function(data) {
                    if (!data || data.length < 100) {
                        reject(new Error("GeoGebra returnerede et tomt billede \u2014 tegn noget f\u00f8rst"));
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
            range.insertInlinePictureFromBase64(base64, "End");
            await window.ensureFooter(context);
            await context.sync();
        });

        setModalStatus("\u2705 Graf indsat i Word!");
        // Rul status tilbage til normal efter 3 sekunder
        setTimeout(() => setModalStatus("Klar \u2014 tegn en funktion og tryk \u00bbInds\u00e6t i Word\u00ab"), 3000);

    } catch (e) {
        console.error("Fejl ved indsætning af graf:", e);
        setModalStatus("\u274c Fejl: " + e.message);
    }
}

// ==========================================
// EVENT LISTENERS
// ==========================================
document.addEventListener('DOMContentLoaded', function() {

    // --- Graf-tab knapper ---
    const openModalBtn = document.getElementById('ggbOpenModalBtn');
    if (openModalBtn) {
        openModalBtn.onclick = function() {
            openGgbModal();
        };
    }

    // "Hent fra Word" i Graf-tab (åbner modal hvis nødvendigt)
    const loadEqBtn = document.getElementById('ggbLoadEquationBtn');
    if (loadEqBtn) {
        loadEqBtn.onclick = async function() {
            openGgbModal();
            // Vent på at GeoGebra er klar (max 10 sekunder)
            if (!window.ggbApplet) {
                let waited = 0;
                await new Promise(resolve => {
                    const check = setInterval(() => {
                        waited += 100;
                        if (window.ggbApplet || waited > 10000) {
                            clearInterval(check);
                            resolve();
                        }
                    }, 100);
                });
            }
            await loadEquationFromWord();
        };
    }

    // --- Modal knapper ---
    const modalLoadBtn = document.getElementById('ggbModalLoadBtn');
    if (modalLoadBtn) {
        modalLoadBtn.onclick = loadEquationFromWord;
    }

    const insertBtn = document.getElementById('ggbInsertBtn');
    if (insertBtn) {
        insertBtn.onclick = insertGraphToWord;
    }

    const closeModalBtn = document.getElementById('ggbCloseModalBtn');
    if (closeModalBtn) {
        closeModalBtn.onclick = closeGgbModal;
    }

    // A4: Escape-keydown lytter globalt, men vi tjekker nu at modalen rent faktisk
    // er åben inden vi lukker den — undgår at forstyrre andre tastatur-shortcuts.
    document.addEventListener('keydown', function(e) {
        if (e.key === 'Escape') {
            const modal = document.getElementById('ggbModal');
            if (modal && modal.classList.contains('open')) closeGgbModal();
        }
    });
});

// Gør openGgbModal tilgængelig globalt så regression.js kan åbne modalen
window.openGgbModal  = openGgbModal;
window.closeGgbModal = closeGgbModal;
