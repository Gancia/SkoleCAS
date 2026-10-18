// ==========================================
// HOVEDMODUL – UI-INITIALISERING & WORD HJÆLPERE
// ==========================================

window.currentResult = "";
window.isOfficeReady = false;
let lastFocusedPanelInput = null;

if (typeof Office !== "undefined" && Office.onReady) {
    Office.onReady(function(info) {
        if (info && info.host === Office.HostType.Word) {
            window.isOfficeReady = true;
        }
    });
}

/**
 * Skifter aktiv fane og opdaterer ARIA-attributter.
 */
window.switchTab = function(tabId, btnElement) {
    const tabs = document.querySelectorAll('.tab-content');
    for (let i = 0; i < tabs.length; i++) {
        tabs[i].classList.remove('active');
        tabs[i].setAttribute('hidden', '');
    }

    const btns = document.querySelectorAll('.tab-btn');
    for (let j = 0; j < btns.length; j++) {
        btns[j].classList.remove('active');
        btns[j].setAttribute('aria-selected', 'false');
    }

    const activeTab = document.getElementById('tab-' + tabId);
    if (activeTab) {
        activeTab.classList.add('active');
        activeTab.removeAttribute('hidden');
    }

    const targetBtn = btnElement || document.getElementById('tabBtn-' + tabId);
    if (targetBtn) {
        targetBtn.classList.add('active');
        targetBtn.setAttribute('aria-selected', 'true');
    }
};

/**
 * Indsætter et matematisk symbol enten i panelets inputfelt eller ved markøren i Word.
 */
window.insertSymbol = async function(symbol) {
    const errorBox = document.getElementById("inlineError");
    if (errorBox) errorBox.textContent = "";

    const inputEl = document.getElementById("mathInput");

    if (inputEl) {
        const start = typeof inputEl.selectionStart === "number" ? inputEl.selectionStart : inputEl.value.length;
        const end = typeof inputEl.selectionEnd === "number" ? inputEl.selectionEnd : inputEl.value.length;
        const val = inputEl.value;
        inputEl.value = val.slice(0, start) + symbol + val.slice(end);

        // Hvis symbolet ender på "()", placér markøren inde i parentesen
        const newPos = symbol.endsWith("()") ? start + symbol.length - 1 : start + symbol.length;
        inputEl.focus();
        inputEl.setSelectionRange(newPos, newPos);
        inputEl.dispatchEvent(new Event("input", { bubbles: true }));
        return;
    }
};

/**
 * Sikrer at dokumentets sidefod indeholder "Løst med Skole CAS".
 */
window.ensureFooter = async function(context) {
    try {
        const section = context.document.sections.getFirst();
        const footer = section.getFooter("Primary");
        footer.load("text");
        await context.sync();

        if (!footer.text || !footer.text.includes("Løst med Skole CAS")) {
            const p = footer.insertParagraph("Løst med Skole CAS", "End");
            p.font.color = "#8a8886";
            p.font.size = 10;
            p.alignment = "Right";
            await context.sync();
        }
    } catch (e) {
        console.warn("Kunne ikke opdatere sidefod:", e);
    }
};

/**
 * Sletter eventuelt låste Content Controls eller markering i Word.
 */
window.clearWordSelection = async function() {
    const errorBox = document.getElementById("inlineError");
    if (errorBox) errorBox.textContent = "";

    if (!window.isOfficeReady || !window.Word) {
        if (errorBox) errorBox.textContent = "Word-forbindelsen er ikke aktiv.";
        return;
    }

    try {
        await Word.run(async (context) => {
            const range = context.document.getSelection();
            const parentCc = range.parentContentControlOrNullObject;
            parentCc.load("isNullObject");

            const ccs = range.contentControls;
            ccs.load("items");
            await context.sync();

            if (!parentCc.isNullObject) {
                parentCc.cannotDelete = false;
                parentCc.cannotEdit = false;
                parentCc.delete(false);
            } else if (ccs.items.length > 0) {
                for (let i = 0; i < ccs.items.length; i++) {
                    ccs.items[i].cannotDelete = false;
                    ccs.items[i].cannotEdit = false;
                    ccs.items[i].delete(false);
                }
            } else {
                range.clear();
            }

            await context.sync();
        });
    } catch (error) {
        console.error("Fejl ved sletning af markering:", error);
        if (errorBox) errorBox.textContent = "Kunne ikke rydde markeringen i Word.";
    }
};

/**
 * Tema-styring (Lys / Mørk) med sikker localStorage-håndtering.
 */
function applyTheme(theme) {
    const resolved = theme === "dark" ? "dark" : "light";
    document.documentElement.setAttribute("data-theme", resolved);
    const toggleBtn = document.getElementById("themeToggle");
    if (toggleBtn) {
        toggleBtn.setAttribute(
            "aria-label",
            resolved === "dark" ? "Skift til lyst tema" : "Skift til mørkt tema"
        );
        toggleBtn.setAttribute(
            "title",
            resolved === "dark" ? "Skift til lyst tema" : "Skift til mørkt tema"
        );
    }
}

window.toggleTheme = function() {
    const currentTheme = document.documentElement.getAttribute("data-theme");
    const newTheme = currentTheme === "dark" ? "light" : "dark";
    applyTheme(newTheme);
    try {
        localStorage.setItem("theme", newTheme);
    } catch (e) {
        // Ignorér hvis localStorage er blokeret i iframe/privat tilstand
    }
};

document.addEventListener("DOMContentLoaded", function() {
    // Indlæs gemt tema
    try {
        const savedTheme = localStorage.getItem("theme");
        if (savedTheme) {
            applyTheme(savedTheme);
        } else {
            applyTheme("light");
        }
    } catch (e) {
        applyTheme("light");
    }

    const themeToggle = document.getElementById("themeToggle");
    if (themeToggle) {
        themeToggle.addEventListener("click", window.toggleTheme);
    }

    // Faneblade
    const tabBtns = document.querySelectorAll('.tab-btn');
    for (let i = 0; i < tabBtns.length; i++) {
        tabBtns[i].addEventListener('click', function() {
            window.switchTab(this.getAttribute('data-tab'), this);
        });
    }

    // Husk fokus på panelets matematikfelt for symbolknapper
    const mathInput = document.getElementById("mathInput");
    if (mathInput) {
        mathInput.addEventListener("focus", function() {
            lastFocusedPanelInput = mathInput;
        });
    }

    // Symbolknapper (data-symbol attribut)
    const symbolBtns = document.querySelectorAll('[data-symbol]');
    for (let i = 0; i < symbolBtns.length; i++) {
        symbolBtns[i].addEventListener("mousedown", function(e) {
            // Forhindr at klikket fjerner fokus fra mathInput
            e.preventDefault();
        });
        symbolBtns[i].addEventListener("click", function() {
            window.insertSymbol(this.getAttribute("data-symbol"));
        });
    }

    // Ryd markering i Word
    const deleteMathBtn = document.getElementById("deleteMathBtn");
    if (deleteMathBtn) {
        deleteMathBtn.addEventListener('click', window.clearWordSelection);
    }

    // Beregner-knapper
    const bindClick = (id, fn) => {
        const el = document.getElementById(id);
        if (el) el.addEventListener("click", fn);
    };

    bindClick("inlineCalcBtn", () => { if (window.calculateInline) window.calculateInline('eval'); });
    bindClick("inlineSolveBtn", () => { if (window.calculateInline) window.calculateInline('solve'); });
    bindClick("inlineDiffBtn", () => { if (window.calculateInline) window.calculateInline('diff'); });
    bindClick("inlineIntBtn", () => { if (window.calculateInline) window.calculateInline('int'); });
    bindClick("fetchFromWordBtn", () => { if (window.fetchFromWordToPanel) window.fetchFromWordToPanel(); });
    bindClick("clearMathInputBtn", () => {
        if (mathInput) {
            mathInput.value = "";
            if (window.updateMathPreview) window.updateMathPreview();
            mathInput.focus();
        }
        const err = document.getElementById("inlineError");
        if (err) err.textContent = "";
    });
    bindClick("insertResultToWordBtn", () => {
        if (window.insertCalcObjectToWord) window.insertCalcObjectToWord(window.lastCalculation);
    });

    // Indstillinger i Beregner
    const exactToggle = document.getElementById("exactToggle");
    const decimalOptions = document.getElementById("decimalOptions");
    const stackedToggle = document.getElementById("stackedFractionToggle");
    const decimalSelect = document.getElementById("decimalSelect");
    const separatorSelect = document.getElementById("decimalSeparator");

    if (exactToggle && decimalOptions) {
        exactToggle.addEventListener("change", function() {
            decimalOptions.style.display = this.checked ? "none" : "flex";
            if (window.updateMathPreview) window.updateMathPreview();
        });
    }
    if (stackedToggle) {
        stackedToggle.addEventListener("change", () => {
            if (window.updateMathPreview) window.updateMathPreview();
        });
    }
    if (decimalSelect) {
        decimalSelect.addEventListener("change", () => {
            if (window.updateMathPreview) window.updateMathPreview();
        });
    }
    if (separatorSelect) {
        separatorSelect.addEventListener("change", () => {
            if (window.updateMathPreview) window.updateMathPreview();
        });
    }

    // Trekantsberegner
    bindClick("solveTriangleBtn", () => { if (window.solveTriangle) window.solveTriangle(); });
    bindClick("resetTriangleBtn", () => { if (window.resetTriangle) window.resetTriangle(); });

    // GeoGebra Online-bro (konverterer til punktum, kopierer og åbner geogebra.org)
    bindClick("openGeoGebraBtn", () => {
        if (window.GeoGebraBridge) window.GeoGebraBridge.openOnlineWithExpression();
    });

    // Tillad Enter i trekant-felterne til at starte beregning
    const triInputs = document.querySelectorAll('#tab-trekant input');
    for (let i = 0; i < triInputs.length; i++) {
        triInputs[i].addEventListener("keydown", function(e) {
            if (e.key === "Enter") {
                e.preventDefault();
                if (window.solveTriangle) window.solveTriangle();
            }
        });
    }
});
