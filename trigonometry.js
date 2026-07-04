// ==========================================
// TREKANTSBEREGNER – med fuld SSA-håndtering
// ==========================================

/**
 * Forsøger at fuldende en trekant fra delvist kendte værdier.
 * Kører iterativt med sinusrelationen og cosinusrelationen.
 *
 * @param {object}  tri               - {a,b,c,A,B,C} – null for ukendte værdier
 * @param {boolean} useObtuseForSines - Brug 180°-arcsin (stump vinkel) i stedet for arcsin.
 *                                      Aktiveres kun ved det andet kald i SSA-tilfældet.
 * @param {number}  DEG               - Math.PI / 180
 * @returns {object|null}  Udfyldt {a,b,c,A,B,C} eller null hvis trekanten er umulig
 */
function tryCompleteTri(tri, useObtuseForSines, DEG) {
    let { a, b, c, A, B, C } = tri;

    for (let iter = 0; iter < 12; iter++) {

        // 1. Vinkelsummen i en trekant = 180°
        if (A !== null && B !== null && C === null) C = 180 - A - B;
        if (A !== null && C !== null && B === null) B = 180 - A - C;
        if (B !== null && C !== null && A === null) A = 180 - B - C;

        // 2. Sinusrelationen: side/sin(vinkel) = konstant (k)
        //    Find det FØRSTE kendte (side, vinkel)-par og brug det som reference.
        const pairs = [[a, A], [b, B], [c, C]];
        const knownPair = pairs.find(([s, ang]) => s !== null && ang !== null);

        if (knownPair) {
            const [ks, kang] = knownPair;
            const k = ks / Math.sin(kang * DEG);

            // Beregn ukendte SIDER fra kendte vinkler (direkte multiplikation, ingen fejl)
            if (a === null && A !== null) a = k * Math.sin(A * DEG);
            if (b === null && B !== null) b = k * Math.sin(B * DEG);
            if (c === null && C !== null) c = k * Math.sin(C * DEG);

            // Hjælpefunktion: beregn en vinkel fra en side via sinusrelationen.
            // useObtuseForSines styrer om vi vælger den akutte eller stumpe løsning.
            // Det er netop her SSA-tvetydigheden opstår: sin(B) = sin(180°-B).
            const computeAngle = (side) => {
                const sinVal = side / k;
                if (sinVal > 1 + 1e-9) return null; // Umulig trekant (siden for lang)
                const acute = Math.asin(Math.min(sinVal, 1)) / DEG;
                return useObtuseForSines ? (180 - acute) : acute;
            };

            // Beregn ukendte VINKLER fra kendte sider
            if (A === null && a !== null) A = computeAngle(a);
            if (B === null && b !== null) B = computeAngle(b);
            if (C === null && c !== null) C = computeAngle(c);
        }

        // 3. Cosinusrelationen: a² = b² + c² − 2bc·cos(A)  (ingen tvetydighed)
        if (b !== null && c !== null && A !== null && a === null)
            a = Math.sqrt(b*b + c*c - 2*b*c * Math.cos(A * DEG));
        if (a !== null && c !== null && B !== null && b === null)
            b = Math.sqrt(a*a + c*c - 2*a*c * Math.cos(B * DEG));
        if (a !== null && b !== null && C !== null && c === null)
            c = Math.sqrt(a*a + b*b - 2*a*b * Math.cos(C * DEG));

        // 4. Omvendt cosinusrelation: find VINKLER når alle tre sider er kendte (SSS)
        //    acos giver altid ét entydigt resultat i [0°, 180°].
        if (a !== null && b !== null && c !== null) {
            const cosA = (b*b + c*c - a*a) / (2*b*c);
            const cosB = (a*a + c*c - b*b) / (2*a*c);
            const cosC = (a*a + b*b - c*c) / (2*a*b);
            if (A === null && Math.abs(cosA) <= 1) A = Math.acos(cosA) / DEG;
            if (B === null && Math.abs(cosB) <= 1) B = Math.acos(cosB) / DEG;
            if (C === null && Math.abs(cosC) <= 1) C = Math.acos(cosC) / DEG;
        }
    }

    // --- Valider løsningen ---
    if ([a, b, c, A, B, C].some(x => x === null || !isFinite(x) || isNaN(x))) return null;
    if (a <= 0 || b <= 0 || c <= 0)   return null; // Negative/nul sider er umulige
    if (A <= 0 || B <= 0 || C <= 0)   return null; // Negative/nul vinkler er umulige
    if (Math.abs(A + B + C - 180) > 0.01) return null; // Vinkelsummen skal være ~180°

    return { a, b, c, A, B, C };
}

/**
 * Afgør om de givne 3 kendte værdier udgør et SSA-tilfælde (tvetydigt).
 *
 * SSA = to sider + én IKKE-indesluttet vinkel (modstående vinkel).
 * SAS = to sider + den INDESLUTTEDE vinkel (imellem de to sider).
 *
 * I standardnotation er vinkel X indesluttet mellem de to sider
 * der støder op til hjørne X:
 *   - Vinkel A er indesluttet mellem siderne b og c
 *   - Vinkel B er indesluttet mellem siderne a og c
 *   - Vinkel C er indesluttet mellem siderne a og b
 *
 * @returns {boolean}
 */
function isSSA(a, b, c, A, B, C) {
    const sides  = [a, b, c].filter(x => x !== null).length;
    const angles = [A, B, C].filter(x => x !== null).length;
    if (sides !== 2 || angles !== 1) return false;

    // Disse tre tilfælde er SAS (indesluttet vinkel) → ikke SSA
    if (a !== null && b !== null && C !== null) return false; // C er indesluttet mellem a og b
    if (a !== null && c !== null && B !== null) return false; // B er indesluttet mellem a og c
    if (b !== null && c !== null && A !== null) return false; // A er indesluttet mellem b og c

    // Alt andet med 2 sider og 1 vinkel er SSA (modstående vinkel)
    return true;
}

/**
 * Afgør om to trekantløsninger er matematisk identiske.
 */
function isSameSolution(sol1, sol2) {
    if (!sol1 || !sol2) return false;
    // A12: 0.01° (36 buesekunder) var for grov og kunne fejlklassificere
    // to næsten identiske SSA-løsninger som én løsning.
    // 1e-4° ≈ 0.36 buesekunder er matematisk præcis og stadig robust.
    return Math.abs(sol1.A - sol2.A) < 1e-4 &&
           Math.abs(sol1.B - sol2.B) < 1e-4 &&
           Math.abs(sol1.C - sol2.C) < 1e-4;
}

// ==========================================
// HOVED-FUNKTION
// ==========================================
window.solveTriangle = function() {
    const errorBox = document.getElementById("triError");
    errorBox.innerHTML = "";

    const DEG = Math.PI / 180;
    const readNum = (id) => {
        const v = parseFloat(document.getElementById(id).value);
        return isNaN(v) ? null : v;
    };

    let a = readNum("tri_a"), b = readNum("tri_b"), c = readNum("tri_c");
    let A = readNum("tri_A"), B = readNum("tri_B"), C = readNum("tri_C");

    // --- Inputvalidering ---
    const sideEntries  = [{ val: a, name: 'a' }, { val: b, name: 'b' }, { val: c, name: 'c' }];
    const angleEntries = [{ val: A, name: 'A' }, { val: B, name: 'B' }, { val: C, name: 'C' }];

    for (const { val, name } of sideEntries) {
        if (val !== null && val <= 0) {
            errorBox.innerText = `Side ${name} skal være et positivt tal (du indtastede ${val}).`;
            return;
        }
    }
    for (const { val, name } of angleEntries) {
        if (val !== null && (val <= 0 || val >= 180)) {
            errorBox.innerText = `Vinkel ${name} skal ligge strengt mellem 0° og 180° (du indtastede ${val}°).`;
            return;
        }
    }

    const knownCount = [a, b, c, A, B, C].filter(x => x !== null).length;
    if (knownCount < 3) {
        errorBox.innerText = "Du skal indtaste præcis 3 kendte værdier.";
        return;
    }
    if (knownCount > 3) {
        errorBox.innerText = "Du har indtastet mere end 3 værdier — slet de overflødige.";
        return;
    }
    if (a === null && b === null && c === null) {
        errorBox.innerText = "Du skal opgive mindst én side — tre vinkler alene bestemmer ikke trekanten.";
        return;
    }

    // Tjek om de kendte vinkler allerede summer til ≥ 180°
    const knownAngles = [A, B, C].filter(x => x !== null);
    if (knownAngles.length >= 2) {
        const angSum = knownAngles.reduce((s, v) => s + v, 0);
        if (angSum >= 180) {
            errorBox.innerText = `Vinklerne ${knownAngles.join('° + ')}° = ${angSum}° ≥ 180° — umulig trekant.`;
            return;
        }
    }

    // --- Løsning ---
    const sol1 = tryCompleteTri({ a, b, c, A, B, C }, false, DEG);

    if (!sol1) {
        errorBox.innerText = "Ingen gyldig trekant med disse værdier. Kontrollér at tallene passer sammen.";
        return;
    }

    const solutions = [sol1];

    // Tjek for SSA-tvetydighed (den "dumme side" – to mulige løsninger)
    if (isSSA(a, b, c, A, B, C)) {
        const sol2 = tryCompleteTri({ a, b, c, A, B, C }, true, DEG);
        if (sol2 && !isSameSolution(sol1, sol2)) {
            solutions.push(sol2);
            errorBox.innerHTML =
                "<strong>⚠️ Tvetydigt SSA-tilfælde:</strong> Der eksisterer " +
                "<strong>to gyldige trekanter</strong> med disse mål! " +
                "Begge løsninger er indsat i dit dokument.";
        }
    }

    solutions.forEach((sol, idx) => {
        const label = solutions.length > 1 ? `Løsning ${idx + 1}` : null;
        window.insertTriangleToWord(sol, label);
    });
};

// ==========================================
// INDSÆT RESULTAT I WORD
// ==========================================
window.insertTriangleToWord = async function(data, label) {
    // Formater ét tal: sikker afrunding + dansk komma + fjern trailing nuls
    const fmt = (n) => {
        const fn = window.safeRound ? window.safeRound : (v, d) => Math.round(v * Math.pow(10, d)) / Math.pow(10, d);
        const r = fn(n, 4); // Behold op til 4 decimaler i mellemregning
        return r.toFixed(2)
                .replace(/(\.\d*[1-9])0+$/, '$1')  // Fjern trailing nuls (2.50 → 2.5)
                .replace(/\.0*$/, '')               // Fjern tom decimal (2. → 2)
                .replace('.', ',');                 // Dansk komma
    };

    try {
        await Word.run(async (context) => {
            const range = context.document.getSelection();

            const title = label ? `Trekantsberegning (${label}):` : "Trekantsberegning:";
            let header = range.insertParagraph(title, Word.InsertLocation.after);
            header.font.bold  = true;
            header.font.color = "#0078d4";

            let content1 = header.insertParagraph(
                `Sider:   a = ${fmt(data.a)}  |  b = ${fmt(data.b)}  |  c = ${fmt(data.c)}`,
                Word.InsertLocation.after
            );
            content1.font.italic = true;
            content1.font.color  = "#333333";

            let content2 = content1.insertParagraph(
                `Vinkler: A = ${fmt(data.A)}°  |  B = ${fmt(data.B)}°  |  C = ${fmt(data.C)}°`,
                Word.InsertLocation.after
            );
            content2.font.italic = true;
            content2.font.color  = "#333333";

            // Nulstil formatering til normal tekst
            let empty = content2.insertParagraph("", Word.InsertLocation.after);
            empty.font.bold   = false;
            empty.font.italic = false;

            await window.ensureFooter(context);
            await context.sync();
        });
    } catch (e) {
        console.error("Fejl ved indsætning af trekant:", e);
    }
};
