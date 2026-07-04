let currentStatChart = null;

// Hjælpefunktion: formater tal til dansk decimaltegn med safeRound
function statFmt(n, decimals) {
    const d = decimals !== undefined ? decimals : 2;
    const r = window.safeRound ? window.safeRound(n, d) : Math.round(n * Math.pow(10, d)) / Math.pow(10, d);
    return r.toFixed(d).replace('.', ',');
}

function parseData(input) {
    const rawTokens = input.split(/[\s;|\t\n]+/);
    const data = [];
    for (let token of rawTokens) {
        token = token.trim();
        if (token !== '') {
            // Håndter dansk (1.234,56) vs engelsk (1,234.56) format
            if (token.includes('.') && token.includes(',')) {
                if (token.lastIndexOf(',') > token.lastIndexOf('.')) {
                    token = token.replace(/\./g, '').replace(',', '.');
                } else {
                    token = token.replace(/,/g, '');
                }
            } else if (token.includes(',')) {
                token = token.replace(/,/g, '.');
            }
            const num = parseFloat(token);
            if (!isNaN(num)) {
                data.push(num);
            }
        }
    }
    return data;
}

/**
 * Parser grupperede data.
 * Understøtter to formater:
 *   Format A (2 tal pr. linje):  midtpunkt frekvens  →  f.eks. "15 4"
 *   Format B (3 tal pr. linje):  start slut frekvens →  f.eks. "10 20 4"
 */
function parseGroupedData(input) {
    const lines = input.trim().split(/\n/);
    const groups = [];

    for (let line of lines) {
        line = line.trim();
        if (!line) continue;

        // Fjern eventuelle interval-tegn ([, ), (, ]) fra dansk notations
        line = line.replace(/[\[\]()<>]/g, '');
        // Erstat semikolon med mellemrum
        line = line.replace(/;/g, ' ');
        // Konverter dansk decimal-komma ("1,5" → "1.5") men IKKE separatorkomma
        line = line.replace(/(\d),(\d)/g, '$1.$2');
        // Erstat resterende kommaer med mellemrum (bruges som separator)
        line = line.replace(/,/g, ' ');
        // Erstat bindestreg der IKKE starter et negativt tal med mellemrum
        line = line.replace(/(\d)-(\d)/g, '$1 $2');

        const nums = (line.match(/[-]?[\d.]+/g) || []).map(Number).filter(n => !isNaN(n));
        if (nums.length < 2) continue;

        if (nums.length >= 3) {
            // Format B: start, slut, frekvens
            const lower = nums[0], upper = nums[1], freq = nums[2];
            const width = upper - lower;
            const midpoint = (lower + upper) / 2;
            groups.push({ lower, upper, midpoint, freq, width });
        } else {
            // Format A: midtpunkt, frekvens
            groups.push({ lower: null, upper: null, midpoint: nums[0], freq: nums[1], width: null });
        }
    }

    return groups;
}

window.calculateUngrouped = async function() {
    const inputArea = document.getElementById("statInput");
    const errorDiv = document.getElementById("statError");
    if (errorDiv) errorDiv.textContent = "";
    
    if (!inputArea) return;
    
    const data = parseData(inputArea.value);
    
    if (data.length === 0) {
        if (errorDiv) errorDiv.textContent = "Kunne ikke finde nogle gyldige tal.";
        return;
    }
    
    // Sorter data for en sikkerheds skyld
    data.sort((a, b) => a - b);
    
    const n = data.length;
    const min = ss.min(data);
    const max = ss.max(data);
    const mean = ss.mean(data);
    const median = ss.median(data);
    const q1 = ss.quantile(data, 0.25);
    const q3 = ss.quantile(data, 0.75);
    
    // simple-statistics kan fejle ved variance hvis n=1
    let variance = 0;
    let stddev = 0;
    if (n > 1) {
        variance = ss.sampleVariance(data);
        stddev = ss.sampleStandardDeviation(data);
    }
    
    // Create Boxplot
    const canvas = document.getElementById("statChart");
    if (canvas) {
        canvas.style.display = "block";
        const ctx = canvas.getContext("2d");
        
        if (currentStatChart) {
            currentStatChart.destroy();
        }
        
        currentStatChart = new Chart(ctx, {
            type: 'boxplot',
            data: {
                labels: ['Ugrupperede Data'],
                datasets: [{
                    label: 'Datasæt',
                    backgroundColor: 'rgba(54, 162, 235, 0.5)',
                    borderColor: 'rgba(54, 162, 235, 1)',
                    borderWidth: 1,
                    outlierColor: '#999999',
                    padding: 10,
                    itemRadius: 0,
                    data: [data]
                }]
            },
            options: {
                responsive: false,
                animation: false,
                plugins: {
                    legend: {
                        display: false
                    }
                }
            }
        });
    }

    // Vent kort på at grafen tegnes
    await new Promise(r => setTimeout(r, 100));

    let base64Data = null;
    if (currentStatChart) {
        const base64Image = currentStatChart.toBase64Image();
        // Fjern prefixet "data:image/png;base64," som Chart.js genererer, for Word.run forventer kun ren Base64
        base64Data = base64Image.replace(/^data:image\/(png|jpeg);base64,/, "");
    }
    
    // Insert into Word
    try {
        await Word.run(async (context) => {
            const range = context.document.getSelection();
            
            // Format table data
            // VIGTIGT: Vi bruger ss.sampleVariance og ss.sampleStandardDeviation
            // som begge dividerer med (n-1) → det er STIKPRØVESTATISTIKKER (s², s).
            // Brug IKKE σ (sigma) som label – σ refererer til populationsstatistikker (÷n).
            const tableData = [
                ["Mål", "Værdi"],
                ["Antal (n)", n.toString()],
                ["Minimum", statFmt(min)],
                ["1. Kvartil (Q1)", statFmt(q1)],
                ["Median", statFmt(median)],
                ["3. Kvartil (Q3)", statFmt(q3)],
                ["Maksimum", statFmt(max)],
                ["Gennemsnit (x̄)", statFmt(mean)],
                ["Varians (s²)", statFmt(variance)],
                ["Spredning (s)", statFmt(stddev)]
            ];
            
            const table = range.insertTable(tableData.length, 2, "After", tableData);
            table.style = "Grid Table 4 - Accent 1";
            
            if (base64Data) {
                // Indsæt billedet efter tabellen
                const paragraph = table.getRange().insertParagraph("", "After");
                paragraph.insertInlinePictureFromBase64(base64Data);
            }
            
            await window.ensureFooter(context);
            await context.sync();
        });
    } catch (error) {
        console.error(error);
        if (errorDiv) errorDiv.textContent = "Fejl ved indsættelse i Word: " + error.message;
    }
}

window.calculateGrouped = async function() {
    const inputArea = document.getElementById("statInput");
    const errorDiv  = document.getElementById("statError");
    if (errorDiv) errorDiv.textContent = "";
    if (!inputArea) return;

    const groups = parseGroupedData(inputArea.value);

    if (groups.length === 0) {
        if (errorDiv) errorDiv.textContent =
            "Format: Én klasse pr. linje med midtpunkt og frekvens: \"15 4\"  " +
            "(eller start slut frekvens: \"10 20 4\")";
        return;
    }
    if (groups.some(g => !isFinite(g.freq) || g.freq <= 0)) {
        if (errorDiv) errorDiv.textContent = "Alle frekvenser skal være positive tal.";
        return;
    }

    const n    = groups.reduce((s, g) => s + g.freq, 0);
    const mean = groups.reduce((s, g) => s + g.midpoint * g.freq, 0) / n;

    // Stikprøvevarians: s² = Σ(fi · (xi − x̄)²) / (n−1)
    const variance = n > 1
        ? groups.reduce((s, g) => s + g.freq * Math.pow(g.midpoint - mean, 2), 0) / (n - 1)
        : 0;
    const stddev = Math.sqrt(variance);

    // Beregn kvartiler og median
    // Hvis vi har fulde intervaller → lineær interpolation inden for klasserne (eksakt metode)
    // Ellers → tilnærmelse via sorteret liste af midtpunkter
    const hasIntervals = groups.every(g => g.lower !== null && g.width !== null && g.width > 0);
    let median, q1, q3, min, max;

    if (hasIntervals) {
        min = groups[0].lower;
        max = groups[groups.length - 1].upper;

        // Lineær interpolation: Median/Q1/Q3 beregnes som L + ((target − F_cum) / f) * h
        const interpolate = (targetFreq) => {
            let cumFreq = 0;
            for (const g of groups) {
                const prevCum = cumFreq;
                cumFreq += g.freq;
                if (cumFreq >= targetFreq) {
                    return g.lower + ((targetFreq - prevCum) / g.freq) * g.width;
                }
            }
            return max;
        };

        median = interpolate(n / 2);
        q1     = interpolate(n / 4);
        q3     = interpolate(3 * n / 4);
    } else {
        // Midpunkts-baseret tilnærmelse: byg en udvidet liste
        const expanded = [];
        for (const g of groups) {
            for (let i = 0; i < g.freq; i++) expanded.push(g.midpoint);
        }
        expanded.sort((x, y) => x - y);
        min = expanded[0];
        max = expanded[expanded.length - 1];

        const quantile = (q) => {
            const pos = q * (expanded.length - 1);
            const lo = Math.floor(pos), hi = Math.ceil(pos);
            return expanded[lo] + (pos - lo) * (expanded[hi] - expanded[lo]);
        };

        median = quantile(0.5);
        q1     = quantile(0.25);
        q3     = quantile(0.75);
    }

    try {
        await Word.run(async (context) => {
            const range = context.document.getSelection();

            const tableData = [
                ["Mål", "Værdi"],
                ["Antal observationer (n)", n.toString()],
                ["Antal klasser", groups.length.toString()],
                [hasIntervals ? "Mindste klassegrænse" : "Mindste midtpunkt", statFmt(min)],
                ["1. Kvartil (Q1)", statFmt(q1)],
                ["Median", statFmt(median)],
                ["3. Kvartil (Q3)", statFmt(q3)],
                [hasIntervals ? "Største klassegrænse" : "Største midtpunkt", statFmt(max)],
                ["Gennemsnit (x̄)", statFmt(mean)],
                ["Varians (s²)", statFmt(variance)],
                ["Spredning (s)", statFmt(stddev)]
            ];

            const table = range.insertTable(tableData.length, 2, "After", tableData);
            table.style = "Grid Table 4 - Accent 1";

            await window.ensureFooter(context);
            await context.sync();
        });
    } catch (error) {
        console.error(error);
        if (errorDiv) errorDiv.textContent = "Fejl ved indsættelse i Word: " + error.message;
    }
};
