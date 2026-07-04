window.currentResult = "";

window.switchTab = function(tabId, btnElement) {
    // Fjern active class fra alle tab-indhold
    var tabs = document.querySelectorAll('.tab-content');
    for (var i = 0; i < tabs.length; i++) {
        tabs[i].classList.remove('active');
    }
    
    // Fjern active class fra alle knapper
    var btns = document.querySelectorAll('.tab-btn');
    for (var j = 0; j < btns.length; j++) {
        btns[j].classList.remove('active');
    }
    
    // Tilføj active til valgt tab
    var activeTab = document.getElementById('tab-' + tabId);
    if (activeTab) activeTab.classList.add('active');
    
    // Find og aktiver knappen der hører til
    if (btnElement) {
        btnElement.classList.add('active');
    } else {
        var btn = document.getElementById('tabBtn-' + tabId);
        if (btn) btn.classList.add('active');
    }
};

document.addEventListener("DOMContentLoaded", function() {
    var tabBtns = document.querySelectorAll('.tab-btn');
    for (var i = 0; i < tabBtns.length; i++) {
        tabBtns[i].addEventListener('click', function() {
            window.switchTab(this.getAttribute('data-tab'), this);
        });
    }

    // Inline / Dokument knapper
    const inlineCalcBtn = document.getElementById("inlineCalcBtn");
    const exactToggle = document.getElementById("exactToggle");
    const decimalOptions = document.getElementById("decimalOptions");

    if (inlineCalcBtn) {
        inlineCalcBtn.onclick = function() {
            const isExact = exactToggle ? exactToggle.checked : false;
            // O8: Brug window.calculateInline konsekvent frem for bare 'calculateInline'
            if (window.calculateInline) window.calculateInline('eval', isExact);
        };
    }

    if (exactToggle && decimalOptions) {
        exactToggle.onchange = function() {
            decimalOptions.style.display = this.checked ? "none" : "flex";
        };
    }

    if (document.getElementById("inlineSolveBtn")) document.getElementById("inlineSolveBtn").onclick = function() { if (window.calculateInline) window.calculateInline('solve'); };
    if (document.getElementById("inlineDiffBtn"))  document.getElementById("inlineDiffBtn").onclick  = function() { if (window.calculateInline) window.calculateInline('diff'); };
    if (document.getElementById("inlineIntBtn"))   document.getElementById("inlineIntBtn").onclick   = function() { if (window.calculateInline) window.calculateInline('int'); };
    
    // Trekantsberegner
    if (document.getElementById("solveTriangleBtn")) document.getElementById("solveTriangleBtn").onclick = function() { if (window.solveTriangle) window.solveTriangle(); };

    // Kladdehæfte knapper
    if (document.getElementById("calculateBtn")) document.getElementById("calculateBtn").onclick = function() { if (window.calculateDraft) window.calculateDraft(); };
    if (document.getElementById("insertBtn"))    document.getElementById("insertBtn").onclick    = function() { if (window.insertToWord) window.insertToWord(); };

    // Statistik knapper
    const statUngroupedBtn = document.getElementById("statUngroupedBtn");
    if (statUngroupedBtn) {
        statUngroupedBtn.onclick = function() { 
            if (window.calculateUngrouped) window.calculateUngrouped(); 
        };
    }
    const statGroupedBtn = document.getElementById("statGroupedBtn");
    if (statGroupedBtn) {
        statGroupedBtn.onclick = function() {
            if (window.calculateGrouped) window.calculateGrouped();
        };
    }

    // Regression knapper
    const regCalcBtn = document.getElementById("regCalcBtn");
    if (regCalcBtn) {
        regCalcBtn.onclick = function() { 
            if (window.calculateRegression) window.calculateRegression(); 
        };
    }
});

Office.onReady(function(info) {
    if (info.host === Office.HostType.Word) {
        // Office is ready, additional Word-specific setup can go here
    }
});

window.insertSymbol = async function(symbol) {
    try {
        await Word.run(async (context) => {
            const range = context.document.getSelection();
            range.insertText(symbol, "End");
            await context.sync();
        });
    } catch (error) {
        // A2: Vis fejl i UI frem for kun at logge til konsollen
        console.error("Fejl ved indsæt symbol:", error);
        const errorBox = document.getElementById("inlineError");
        if (errorBox) errorBox.innerText = "Kunne ikke indsætte symbol — er Word klar?";
    }
};

window.ensureFooter = async function(context) {
    // Tjekker ALTID sidefod-indholdet i Word frem for at bruge et globalt flag.
    // Det gamle flag (hasAddedFooter) blev aldrig nulstillet, så sidefoden
    // manglede hvis brugeren åbnede et nyt dokument i samme session.
    try {
        const section = context.document.sections.getFirst();
        const footer = section.getFooter("Primary");
        footer.load("text");
        await context.sync();

        if (!footer.text.includes("Løst med Skole CAS")) {
            let p = footer.insertParagraph("Løst med Skole CAS", "End");
            p.font.color = "#8a8886";
            p.font.size = 10;
            p.alignment = "Right";
            await context.sync();
        }
    } catch (e) {
        console.error("Fejl i ensureFooter:", e);
    }
}

window.insertToWord = async function() {
    if (!window.currentResult) return;
    const errorBox = document.getElementById("inlineError");
    try {
        await Word.run(async (context) => {
            const range = context.document.getSelection();
            range.insertText(window.currentResult, "End");
            await window.ensureFooter(context);
            await context.sync();
        });
    } catch (error) {
        // A2: Vis fejl i UI
        console.error(error);
        if (errorBox) errorBox.innerText = "Kunne ikke indsætte i Word — prøv at placere markøren i dokumentet først.";
    }
}
