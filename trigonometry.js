// ==========================================
// TREKANTSBEREGNER – med fuld SSA-håndtering & SVG-skitse
// ==========================================

window.lastTriangleSolutions = [];

/**
 * Forsøger at fuldende en trekant fra 3 kendte værdier.
 * Kører iterativt med vinkelsum, sinusrelationen og cosinusrelationen.
 *
 * @param {object}  tri               - {a,b,c,A,B,C} – null for ukendte værdier
 * @param {boolean} useObtuseForSines - Brug 180°-arcsin (stump vinkel) ved sinusrelation
 * @param {number}  DEG               - Math.PI / 180
 * @returns {object|null}  Udfyldt {a,b,c,A,B,C,area,perimeter} eller null hvis umulig
 */
function tryCompleteTri(tri, useObtuseForSines, DEG) {
    let { a, b, c, A, B, C } = tri;
    let impossible = false;

    for (let iter = 0; iter < 8; iter++) {
        if (a !== null && b !== null && c !== null && A !== null && B !== null && C !== null) {
            break;
        }

        // 1. Vinkelsummen i en trekant = 180°
        if (A !== null && B !== null && C === null) C = 180 - A - B;
        if (A !== null && C !== null && B === null) B = 180 - A - C;
        if (B !== null && C !== null && A === null) A = 180 - B - C;

        // 2. Sinusrelationen: side / sin(vinkel) = k
        const pairs = [[a, A], [b, B], [c, C]];
        const knownPair = pairs.find(([s, ang]) => s !== null && ang !== null && ang > 0 && ang < 180);

        if (knownPair) {
            const [ks, kang] = knownPair;
            const sinKang = Math.sin(kang * DEG);
            if (sinKang <= 1e-12) return null;
            const k = ks / sinKang;

            if (a === null && A !== null && A > 0 && A < 180) a = k * Math.sin(A * DEG);
            if (b === null && B !== null && B > 0 && B < 180) b = k * Math.sin(B * DEG);
            if (c === null && C !== null && C > 0 && C < 180) c = k * Math.sin(C * DEG);

            const computeAngle = (side) => {
                const sinVal = side / k;
                if (sinVal > 1 + 1e-9) {
                    impossible = true;
                    return null;
                }
                const clampedSin = Math.max(-1, Math.min(1, sinVal));
                const acute = Math.asin(clampedSin) / DEG;
                return useObtuseForSines ? (180 - acute) : acute;
            };

            if (A === null && a !== null) A = computeAngle(a);
            if (B === null && b !== null) B = computeAngle(b);
            if (C === null && c !== null) C = computeAngle(c);
            if (impossible) return null;
        }

        // 3. Cosinusrelationen: a² = b² + c² − 2bc·cos(A)
        if (b !== null && c !== null && A !== null && a === null) {
            const val = b * b + c * c - 2 * b * c * Math.cos(A * DEG);
            if (val <= 1e-12) return null;
            a = Math.sqrt(val);
        }
        if (a !== null && c !== null && B !== null && b === null) {
            const val = a * a + c * c - 2 * a * c * Math.cos(B * DEG);
            if (val <= 1e-12) return null;
            b = Math.sqrt(val);
        }
        if (a !== null && b !== null && C !== null && c === null) {
            const val = a * a + b * b - 2 * a * b * Math.cos(C * DEG);
            if (val <= 1e-12) return null;
            c = Math.sqrt(val);
        }

        // 4. Omvendt cosinusrelation (SSS) med clamping mod IEEE 754 afrundingsfejl
        if (a !== null && b !== null && c !== null) {
            if (a <= 1e-9 || b <= 1e-9 || c <= 1e-9) return null;
            const cosA = (b * b + c * c - a * a) / (2 * b * c);
            const cosB = (a * a + c * c - b * b) / (2 * a * c);
            const cosC = (a * a + b * b - c * c) / (2 * a * b);

            if (Math.abs(cosA) > 1 + 1e-9 || Math.abs(cosB) > 1 + 1e-9 || Math.abs(cosC) > 1 + 1e-9) {
                return null;
            }
            if (A === null) A = Math.acos(Math.max(-1, Math.min(1, cosA))) / DEG;
            if (B === null) B = Math.acos(Math.max(-1, Math.min(1, cosB))) / DEG;
            if (C === null) C = Math.acos(Math.max(-1, Math.min(1, cosC))) / DEG;
        }
    }

    // Valider løsningen (afvis degenererede og umulige trekanter)
    if ([a, b, c, A, B, C].some(x => x === null || !Number.isFinite(x))) return null;
    if (a <= 1e-7 || b <= 1e-7 || c <= 1e-7) return null;
    if (A <= 1e-5 || B <= 1e-5 || C <= 1e-5) return null;
    if (A >= 180 - 1e-5 || B >= 180 - 1e-5 || C >= 180 - 1e-5) return null;
    if (Math.abs(A + B + C - 180) > 0.01) return null;

    // Trekantsuligheden
    if (a + b <= c + 1e-7 || a + c <= b + 1e-7 || b + c <= a + 1e-7) return null;

    const area = 0.5 * a * b * Math.sin(C * DEG);
    const perimeter = a + b + c;

    return { a, b, c, A, B, C, area, perimeter };
}

/**
 * Afgør om de givne 3 kendte værdier udgør et SSA-tilfælde (tvetydigt).
 */
function isSSA(a, b, c, A, B, C) {
    const sides = [a, b, c].filter(x => x !== null).length;
    const angles = [A, B, C].filter(x => x !== null).length;
    if (sides !== 2 || angles !== 1) return false;

    if (a !== null && b !== null && C !== null) return false;
    if (a !== null && c !== null && B !== null) return false;
    if (b !== null && c !== null && A !== null) return false;

    return true;
}

/**
 * Afgør om to trekantløsninger er matematisk identiske.
 */
function isSameSolution(sol1, sol2) {
    if (!sol1 || !sol2) return false;
    return Math.abs(sol1.A - sol2.A) < 1e-4 &&
           Math.abs(sol1.B - sol2.B) < 1e-4 &&
           Math.abs(sol1.C - sol2.C) < 1e-4;
}

/**
 * Formaterer et tal til dansk visning (op til 2 decimaler, uden overflødige nuller).
 */
function fmtTriNum(n) {
    if (window.formatNumber) {
        return window.formatNumber(n, 2, ',', false);
    }
    const r = Math.round(n * 100) / 100;
    return r.toFixed(2).replace(/(\.\d*[1-9])0+$/, '$1').replace(/\.0*$/, '').replace('.', ',');
}

/**
 * Genererer en målestoksret SVG-tegning af trekanten med hjørner (A, B, C) og sider (a, b, c).
 */
function buildTriangleSvg(sol) {
    const DEG = Math.PI / 180;
    // Placér A i (0, 0), B i (c, 0), og C i (b*cos(A), b*sin(A)) i matematiske koordinater
    const Ax = 0, Ay = 0;
    const Bx = sol.c, By = 0;
    const Cx = sol.b * Math.cos(sol.A * DEG);
    const Cy = sol.b * Math.sin(sol.A * DEG);

    const minX = Math.min(Ax, Bx, Cx);
    const maxX = Math.max(Ax, Bx, Cx);
    const minY = Math.min(Ay, By, Cy);
    const maxY = Math.max(Ay, By, Cy);

    const width = Math.max(maxX - minX, 1e-4);
    const height = Math.max(maxY - minY, 1e-4);

    const svgW = 260;
    const svgH = 145;
    const padX = 36;
    const padY = 26;

    const scale = Math.min((svgW - 2 * padX) / width, (svgH - 2 * padY) / height);
    const drawnW = width * scale;
    const drawnH = height * scale;

    const offsetX = (svgW - drawnW) / 2 - minX * scale;
    const offsetY = (svgH - drawnH) / 2;

    const toSvg = (x, y) => ({
        x: offsetX + x * scale,
        y: svgH - (offsetY + (y - minY) * scale)
    });

    const pA = toSvg(Ax, Ay);
    const pB = toSvg(Bx, By);
    const pC = toSvg(Cx, Cy);

    // Tyngdepunkt (centroid) til at skubbe labels væk fra trekantens indre
    const centX = (pA.x + pB.x + pC.x) / 3;
    const centY = (pA.y + pB.y + pC.y) / 3;

    const pushOut = (pt, dist) => {
        const dx = pt.x - centX;
        const dy = pt.y - centY;
        const len = Math.hypot(dx, dy) || 1;
        return {
            x: pt.x + (dx / len) * dist,
            y: pt.y + (dy / len) * dist
        };
    };

    const lblA = pushOut(pA, 14);
    const lblB = pushOut(pB, 14);
    const lblC = pushOut(pC, 14);

    const midPush = (p1, p2, dist) => {
        const mx = (p1.x + p2.x) / 2;
        const my = (p1.y + p2.y) / 2;
        return pushOut({ x: mx, y: my }, dist);
    };

    const lblSideC = midPush(pA, pB, 12);
    const lblSideA = midPush(pB, pC, 12);
    const lblSideB = midPush(pA, pC, 12);

    return `
    <svg class="tri-svg" viewBox="0 0 ${svgW} ${svgH}" aria-label="Skitse af beregnet trekant">
        <polygon points="${pA.x.toFixed(1)},${pA.y.toFixed(1)} ${pB.x.toFixed(1)},${pB.y.toFixed(1)} ${pC.x.toFixed(1)},${pC.y.toFixed(1)}"
                 class="tri-polygon" />
        <circle cx="${pA.x.toFixed(1)}" cy="${pA.y.toFixed(1)}" r="3" class="tri-vertex" />
        <circle cx="${pB.x.toFixed(1)}" cy="${pB.y.toFixed(1)}" r="3" class="tri-vertex" />
        <circle cx="${pC.x.toFixed(1)}" cy="${pC.y.toFixed(1)}" r="3" class="tri-vertex" />
        <text x="${lblA.x.toFixed(1)}" y="${lblA.y.toFixed(1)}" class="tri-label-angle" text-anchor="middle" dominant-baseline="central">A</text>
        <text x="${lblB.x.toFixed(1)}" y="${lblB.y.toFixed(1)}" class="tri-label-angle" text-anchor="middle" dominant-baseline="central">B</text>
        <text x="${lblC.x.toFixed(1)}" y="${lblC.y.toFixed(1)}" class="tri-label-angle" text-anchor="middle" dominant-baseline="central">C</text>
        <text x="${lblSideA.x.toFixed(1)}" y="${lblSideA.y.toFixed(1)}" class="tri-label-side" text-anchor="middle" dominant-baseline="central">a</text>
        <text x="${lblSideB.x.toFixed(1)}" y="${lblSideB.y.toFixed(1)}" class="tri-label-side" text-anchor="middle" dominant-baseline="central">b</text>
        <text x="${lblSideC.x.toFixed(1)}" y="${lblSideC.y.toFixed(1)}" class="tri-label-side" text-anchor="middle" dominant-baseline="central">c</text>
    </svg>`;
}

/**
 * Viser løsningen (eller begge løsninger ved SSA) i panelet.
 */
function renderTriangleResultsInPanel(solutions, isAmbiguousSSA) {
    window.lastTriangleSolutions = solutions;
    const container = document.getElementById("triResultContainer");
    if (!container) return;

    container.classList.remove("hidden");
    let html = "";

    if (isAmbiguousSSA && solutions.length > 1) {
        html += `
        <div class="tri-warning-banner" role="status">
            <svg class="tri-warning-icon" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
                <path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"></path>
                <line x1="12" y1="9" x2="12" y2="13"></line>
                <line x1="12" y1="17" x2="12.01" y2="17"></line>
            </svg>
            <div class="tri-warning-content">
                <strong>Tvetydigt SSA-tilfælde:</strong> Der findes to gyldige trekanter med disse mål. Begge løsninger er vist herunder.
            </div>
        </div>`;
    }

    solutions.forEach((sol, idx) => {
        const title = solutions.length > 1 ? `Løsning ${idx + 1}` : "Beregnet trekant";
        html += `
        <div class="tri-solution-card">
            <div class="tri-solution-header">
                <h4>${title}</h4>
                <button type="button" class="secondary-btn tri-insert-single-btn" data-sol-idx="${idx}">
                    <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
                        <path d="M5 12h14"></path>
                        <path d="m12 5 7 7-7 7"></path>
                    </svg>
                    <span>Indsæt i Word</span>
                </button>
            </div>
            <div class="tri-svg-wrapper">
                ${buildTriangleSvg(sol)}
            </div>
            <div class="tri-metrics-grid">
                <div class="tri-metric"><span class="tri-metric-label">a</span><span class="tri-metric-val">${fmtTriNum(sol.a)}</span></div>
                <div class="tri-metric"><span class="tri-metric-label">A</span><span class="tri-metric-val">${fmtTriNum(sol.A)}°</span></div>
                <div class="tri-metric"><span class="tri-metric-label">b</span><span class="tri-metric-val">${fmtTriNum(sol.b)}</span></div>
                <div class="tri-metric"><span class="tri-metric-label">B</span><span class="tri-metric-val">${fmtTriNum(sol.B)}°</span></div>
                <div class="tri-metric"><span class="tri-metric-label">c</span><span class="tri-metric-val">${fmtTriNum(sol.c)}</span></div>
                <div class="tri-metric"><span class="tri-metric-label">C</span><span class="tri-metric-val">${fmtTriNum(sol.C)}°</span></div>
            </div>
            <div class="tri-extra-metrics">
                <span><strong>Areal (T):</strong> ${fmtTriNum(sol.area)}</span>
                <span><strong>Omkreds:</strong> ${fmtTriNum(sol.perimeter)}</span>
            </div>
        </div>`;
    });

    container.innerHTML = html;

    const btns = container.querySelectorAll(".tri-insert-single-btn");
    btns.forEach(btn => {
        btn.addEventListener("click", function() {
            const idx = parseInt(this.getAttribute("data-sol-idx"), 10);
            const sol = window.lastTriangleSolutions[idx];
            if (sol) {
                const label = window.lastTriangleSolutions.length > 1 ? `Løsning ${idx + 1}` : null;
                window.insertTriangleToWord(sol, label);
            }
        });
    });
}

// ==========================================
// HOVED-FUNKTION
// ==========================================
window.solveTriangle = function() {
    const errorBox = document.getElementById("triError");
    if (errorBox) errorBox.textContent = "";
    const resultContainer = document.getElementById("triResultContainer");

    const DEG = Math.PI / 180;
    const readNum = (id) => {
        const el = document.getElementById(id);
        if (!el) return null;
        return window.parseDanishNumber ? window.parseDanishNumber(el.value) : null;
    };

    // Tjek først om nogen felter har ugyldige tegn (ikke-tomme men ikke gyldigt tal)
    const fieldDefs = [
        { id: "tri_a", name: "side a" },
        { id: "tri_b", name: "side b" },
        { id: "tri_c", name: "side c" },
        { id: "tri_A", name: "vinkel A" },
        { id: "tri_B", name: "vinkel B" },
        { id: "tri_C", name: "vinkel C" }
    ];

    for (const f of fieldDefs) {
        const el = document.getElementById(f.id);
        if (el && el.value.trim() !== "" && readNum(f.id) === null) {
            if (errorBox) errorBox.textContent = `Ugyldigt tal i ${f.name}. Brug kun cifre og komma/punktum.`;
            return;
        }
    }

    let a = readNum("tri_a"), b = readNum("tri_b"), c = readNum("tri_c");
    let A = readNum("tri_A"), B = readNum("tri_B"), C = readNum("tri_C");

    const sideEntries = [{ val: a, name: 'a' }, { val: b, name: 'b' }, { val: c, name: 'c' }];
    const angleEntries = [{ val: A, name: 'A' }, { val: B, name: 'B' }, { val: C, name: 'C' }];

    for (const { val, name } of sideEntries) {
        if (val !== null && val <= 0) {
            if (errorBox) errorBox.textContent = `Side ${name} skal være et positivt tal (du indtastede ${fmtTriNum(val)}).`;
            return;
        }
    }
    for (const { val, name } of angleEntries) {
        if (val !== null && (val <= 0 || val >= 180)) {
            if (errorBox) errorBox.textContent = `Vinkel ${name} skal ligge mellem 0° og 180° (du indtastede ${fmtTriNum(val)}°).`;
            return;
        }
    }

    const knownCount = [a, b, c, A, B, C].filter(x => x !== null).length;
    if (knownCount < 3) {
        if (errorBox) errorBox.textContent = "Indtast præcis 3 kendte værdier for at beregne trekanten.";
        if (resultContainer) resultContainer.classList.add("hidden");
        return;
    }
    if (knownCount > 3) {
        if (errorBox) errorBox.textContent = "Du har indtastet mere end 3 værdier. Ryd de ekstra felter først.";
        return;
    }
    if (a === null && b === null && c === null) {
        if (errorBox) errorBox.textContent = "Opgiv mindst én side — tre vinkler alene bestemmer ikke trekantens størrelse.";
        return;
    }

    const knownAngles = [A, B, C].filter(x => x !== null);
    if (knownAngles.length >= 2) {
        const angSum = knownAngles.reduce((s, v) => s + v, 0);
        if (angSum >= 180) {
            if (errorBox) errorBox.textContent = `Vinkelsummen af de indtastede vinkler (${fmtTriNum(angSum)}°) er 180° eller derover — umulig trekant.`;
            return;
        }
    }

    const sol1 = tryCompleteTri({ a, b, c, A, B, C }, false, DEG);

    if (!sol1) {
        if (errorBox) errorBox.textContent = "Ingen gyldig trekant med disse værdier. Kontrollér at målene kan danne en trekant.";
        if (resultContainer) resultContainer.classList.add("hidden");
        return;
    }

    const solutions = [sol1];
    let isAmbiguousSSA = false;

    if (isSSA(a, b, c, A, B, C)) {
        const sol2 = tryCompleteTri({ a, b, c, A, B, C }, true, DEG);
        if (sol2 && !isSameSolution(sol1, sol2)) {
            solutions.push(sol2);
            isAmbiguousSSA = true;
        }
    }

    // Vis resultater og SVG-skitse i panelet med det samme
    renderTriangleResultsInPanel(solutions, isAmbiguousSSA);

    // Indsæt også automatisk i Word hvis vi kører i Word
    if (window.isOfficeReady && window.Word) {
        solutions.forEach((sol, idx) => {
            const label = solutions.length > 1 ? `Løsning ${idx + 1}` : null;
            window.insertTriangleToWord(sol, label);
        });
    }
};

/**
 * Nulstiller alle 6 inputfelter og skjuler resultatboksen.
 */
window.resetTriangle = function() {
    const ids = ["tri_a", "tri_b", "tri_c", "tri_A", "tri_B", "tri_C"];
    ids.forEach(id => {
        const el = document.getElementById(id);
        if (el) el.value = "";
    });
    const errorBox = document.getElementById("triError");
    if (errorBox) errorBox.textContent = "";
    const resultContainer = document.getElementById("triResultContainer");
    if (resultContainer) {
        resultContainer.classList.add("hidden");
        resultContainer.innerHTML = "";
    }
    window.lastTriangleSolutions = [];
};

// ==========================================
// INDSÆT RESULTAT I WORD
// ==========================================
window.insertTriangleToWord = async function(data, label) {
    const errorBox = document.getElementById("triError");

    if (!window.isOfficeReady || !window.Word) {
        if (errorBox) errorBox.textContent = "Word-forbindelsen er ikke aktiv (resultatet vises kun i panelet).";
        return;
    }

    try {
        await Word.run(async (context) => {
            const range = context.document.getSelection();

            const title = label ? `Trekantsberegning (${label}):` : "Trekantsberegning:";
            const header = range.insertParagraph(title, Word.InsertLocation.after);
            header.font.bold = true;
            header.font.color = "#0078d4";

            const content1 = header.insertParagraph(
                `Sider:   a = ${fmtTriNum(data.a)}  |  b = ${fmtTriNum(data.b)}  |  c = ${fmtTriNum(data.c)}`,
                Word.InsertLocation.after
            );
            content1.font.italic = true;
            content1.font.color = "#333333";

            const content2 = content1.insertParagraph(
                `Vinkler: A = ${fmtTriNum(data.A)}°  |  B = ${fmtTriNum(data.B)}°  |  C = ${fmtTriNum(data.C)}°`,
                Word.InsertLocation.after
            );
            content2.font.italic = true;
            content2.font.color = "#333333";

            const content3 = content2.insertParagraph(
                `Areal:   T = ${fmtTriNum(data.area)}  |  Omkreds = ${fmtTriNum(data.perimeter)}`,
                Word.InsertLocation.after
            );
            content3.font.italic = true;
            content3.font.color = "#333333";

            const empty = content3.insertParagraph("", Word.InsertLocation.after);
            empty.font.bold = false;
            empty.font.italic = false;

            await window.ensureFooter(context);
            await context.sync();
        });
    } catch (e) {
        console.error("Fejl ved indsættelse af trekant:", e);
        if (errorBox) errorBox.textContent = "Kunne ikke indsætte trekantsberegning i Word: " + e.message;
    }
};
