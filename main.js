// ==========================================
// HOVEDMODUL – UI-INITIALISERING & WORD HJÆLPERE
// ==========================================

window.isOfficeReady = false;
let footerEnsured = false;

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
 * Indsætter et matematisk symbol i panelets inputfelt.
 * Hvis brugeren har markeret tekst i feltet og klikker f.eks. √(x), pakkes markeringen ind i √(...).
 */
window.insertSymbol = function(symbol) {
    const errorBox = document.getElementById("inlineError");
    if (errorBox) errorBox.textContent = "";

    const inputEl = document.getElementById("mathInput");
    if (!inputEl) return;

    const start = typeof inputEl.selectionStart === "number" ? inputEl.selectionStart : inputEl.value.length;
    const end = typeof inputEl.selectionEnd === "number" ? inputEl.selectionEnd : inputEl.value.length;
    const val = inputEl.value;
    const selectedText = val.slice(start, end);

    let replacement = symbol;
    let newPos = start + symbol.length;

    if (selectedText && symbol.endsWith("()")) {
        const prefix = symbol.slice(0, -1);
        replacement = prefix + selectedText + ")";
        newPos = start + replacement.length;
    } else if (selectedText && symbol.startsWith("^")) {
        const wrapped = selectedText.length > 1 && !(selectedText.startsWith("(") && selectedText.endsWith(")"))
            ? `(${selectedText})`
            : selectedText;
        replacement = wrapped + symbol;
        newPos = start + replacement.length;
    } else if (symbol.endsWith("()")) {
        newPos = start + symbol.length - 1;
    }

    inputEl.value = val.slice(0, start) + replacement + val.slice(end);
    inputEl.focus();
    inputEl.setSelectionRange(newPos, newPos);
    inputEl.dispatchEvent(new Event("input", { bubbles: true }));
};

/**
 * Sikrer at dokumentets sidefod indeholder "Løst med Skole CAS".
 * Cacher status i sessionen, så vi undgår unødig context.sync() ved efterfølgende beregninger.
 */
window.ensureFooter = async function(context) {
    if (footerEnsured) return;
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
        footerEnsured = true;
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

function saveSetting(key, val) {
    try {
        localStorage.setItem("skolecas_" + key, String(val));
    } catch (e) {
        // Ignorér hvis localStorage er blokeret
    }
}

function loadSetting(key) {
    try {
        return localStorage.getItem("skolecas_" + key);
    } catch (e) {
        return null;
    }
}

document.addEventListener("DOMContentLoaded", function() {
    // Indlæs gemt tema
    const savedTheme = loadSetting("theme") || (function() {
        try { return localStorage.getItem("theme"); } catch (e) { return null; }
    })();
    applyTheme(savedTheme || "light");

    const themeToggle = document.getElementById("themeToggle");
    if (themeToggle) {
        themeToggle.addEventListener("click", window.toggleTheme);
    }

    // Faneblade med klik og piletast-navigation
    const tabBtns = Array.from(document.querySelectorAll('.tab-btn'));
    tabBtns.forEach((btn, idx) => {
        btn.addEventListener('click', function() {
            window.switchTab(this.getAttribute('data-tab'), this);
        });
        btn.addEventListener('keydown', function(e) {
            if (e.key === 'ArrowRight' || e.key === 'ArrowLeft') {
                e.preventDefault();
                const nextIdx = e.key === 'ArrowRight'
                    ? (idx + 1) % tabBtns.length
                    : (idx - 1 + tabBtns.length) % tabBtns.length;
                const nextBtn = tabBtns[nextIdx];
                if (nextBtn) {
                    window.switchTab(nextBtn.getAttribute('data-tab'), nextBtn);
                    nextBtn.focus();
                }
            }
        });
    });

    const mathInput = document.getElementById("mathInput");

    // Symbolknapper (data-symbol attribut)
    const symbolBtns = document.querySelectorAll('[data-symbol]');
    for (let i = 0; i < symbolBtns.length; i++) {
        symbolBtns[i].addEventListener("mousedown", function(e) {
            // Forhindr at klikket fjerner fokus/markering fra mathInput
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
        if (window.clearPanelCalculationState) {
            window.clearPanelCalculationState();
        }
        const err = document.getElementById("inlineError");
        if (err) err.textContent = "";
        const resultCard = document.getElementById("calcResultCard");
        if (resultCard) resultCard.classList.add("hidden");
        const insertBtn = document.getElementById("insertResultToWordBtn");
        if (insertBtn) insertBtn.disabled = true;
        const ggbNotice = document.getElementById("ggbCopyNotice");
        if (ggbNotice) ggbNotice.classList.add("hidden");
    });
    bindClick("insertResultToWordBtn", () => {
        if (window.insertCalcObjectToWord) window.insertCalcObjectToWord(window.lastCalculation);
    });

    // Indstillinger i Beregner (med localStorage-persistens og øjeblikkelig genberegning)
    const exactToggle = document.getElementById("exactToggle");
    const decimalOptions = document.getElementById("decimalOptions");
    const stackedToggle = document.getElementById("stackedFractionToggle");
    const autoInsertToggle = document.getElementById("autoInsertToggle");
    const decimalSelect = document.getElementById("decimalSelect");
    const separatorSelect = document.getElementById("decimalSeparator");

    const savedExact = loadSetting("exact");
    if (exactToggle && savedExact !== null) {
        exactToggle.checked = savedExact === "true";
    }
    if (exactToggle && decimalOptions) {
        decimalOptions.style.display = exactToggle.checked ? "none" : "flex";
    }

    const savedStacked = loadSetting("stacked");
    if (stackedToggle && savedStacked !== null) {
        stackedToggle.checked = savedStacked === "true";
    }

    const savedAutoInsert = loadSetting("autoInsert");
    if (autoInsertToggle && savedAutoInsert !== null) {
        autoInsertToggle.checked = savedAutoInsert === "true";
    }

    const savedDecimals = loadSetting("decimals");
    if (decimalSelect && savedDecimals !== null) {
        decimalSelect.value = savedDecimals;
    }

    const savedSep = loadSetting("decSep");
    if (separatorSelect && savedSep !== null) {
        separatorSelect.value = savedSep;
    }

    const onSettingsChanged = () => {
        if (window.updateMathPreview) window.updateMathPreview();
        if (window.refreshLastCalculationForSettings) window.refreshLastCalculationForSettings();
    };

    if (exactToggle && decimalOptions) {
        exactToggle.addEventListener("change", function() {
            decimalOptions.style.display = this.checked ? "none" : "flex";
            saveSetting("exact", this.checked);
            onSettingsChanged();
        });
    }
    if (stackedToggle) {
        stackedToggle.addEventListener("change", function() {
            saveSetting("stacked", this.checked);
            onSettingsChanged();
        });
    }
    if (autoInsertToggle) {
        autoInsertToggle.addEventListener("change", function() {
            saveSetting("autoInsert", this.checked);
        });
    }
    if (decimalSelect) {
        decimalSelect.addEventListener("change", function() {
            saveSetting("decimals", this.value);
            onSettingsChanged();
        });
    }
    if (separatorSelect) {
        separatorSelect.addEventListener("change", function() {
            saveSetting("decSep", this.value);
            onSettingsChanged();
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

