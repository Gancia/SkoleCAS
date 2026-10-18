/**
 * SkoleCAS – Delte hjælpefunktioner
 * Indlæses som det første script for at sikre, at alle hjælpemetoder
 * er tilgængelige på tværs af modulerne.
 */

/**
 * Normaliserer et matematisk udtryk til Nerdamer-kompatibelt ASCII-format.
 * Håndterer græske bogstaver, rodtegn, danske typografiske minus/gange-tegn
 * samt Unicode-potenser (x², x³ osv.) kopieret fra Word.
 *
 * @param {string} expr - Rå matematisk streng fra brugeren eller Word
 * @returns {string}
 */
window.normalizeInput = function(expr) {
    if (expr === null || expr === undefined) return "";
    let s = String(expr).trim();
    if (!s) return "";

    // Ensret typografiske operatorer fra Word/PDF til standard ASCII
    s = s
        .replace(/[−–—]/g, '-')
        .replace(/[·⋅×]/g, '*')
        .replace(/÷/g, '/');

    // Konverter Unicode-eksponenter (f.eks. x² → x^2)
    const superscriptMap = {
        '⁰': '0', '¹': '1', '²': '2', '³': '3', '⁴': '4',
        '⁵': '5', '⁶': '6', '⁷': '7', '⁸': '8', '⁹': '9',
        '⁻': '-'
    };
    s = s.replace(/[⁰¹²³⁴⁵⁶⁷⁸⁹⁻]+/g, function(match) {
        const converted = match.split('').map(ch => superscriptMap[ch] || '').join('');
        return converted ? '^(' + converted + ')' : match;
    });
    // Forenkl ^(2) til ^2 for enkelte cifre
    s = s.replace(/\^\((\d)\)/g, '^$1');

    // Normaliser pi og kvadratrod
    s = s
        .replace(/([0-9a-zA-Z)])\s*[πΠ𝜋]/g, '$1*pi')
        .replace(/[πΠ𝜋]/g, 'pi')
        .replace(/([0-9a-zA-Z)])\s*√\s*\(/g, '$1*sqrt(')
        .replace(/√\s*\(/g, 'sqrt(')
        .replace(/([0-9a-zA-Z)])\s*√\s*(-?[a-zA-Z0-9.,]+)/g, '$1*sqrt($2)')
        .replace(/√\s*(-?[a-zA-Z0-9.,]+)/g, 'sqrt($1)');

    return s;
};

/**
 * Konverterer dansk decimalkomma til punktum mellem cifre (f.eks. "3,14" → "3.14").
 * Kører i en løkke så tal med flere grupper håndteres korrekt.
 *
 * @param {string} expr   - Matematisk udtryk
 * @param {string} decSep - Valgt decimalseparator (',' eller '.')
 * @returns {string}
 */
window.convertDanishDecimalToDot = function(expr, decSep) {
    if (!expr || typeof expr !== 'string') return "";
    if (decSep !== ",") return expr;
    let prev = "";
    let curr = expr;
    while (curr !== prev) {
        prev = curr;
        curr = curr.replace(/(\d),(\d)/g, '$1.$2');
    }
    return curr;
};

/**
 * Parser et tal fra tekstinput (håndterer både dansk komma og engelsk punktum).
 * Returnerer null hvis værdien er tom eller ugyldig.
 *
 * @param {string|number} val
 * @returns {number|null}
 */
window.parseDanishNumber = function(val) {
    if (val === null || val === undefined) return null;
    const str = String(val).trim().replace(/[−–—]/g, '-').replace(/,/g, '.');
    if (!str) return null;
    if (!/^[+-]?(\d+(\.\d*)?|\.\d+)$/.test(str)) return null;
    const num = Number(str);
    return Number.isFinite(num) ? num : null;
};

/**
 * Sikker afrunding med eksponentiel notation, som undgår IEEE 754 floating-point fejl
 * både for små (1.005) og store (1000.005) tal, samt negativ nul (-0).
 *
 * @param {number} num      - Tal der skal afrundes
 * @param {number} decimals - Antal ønskede decimaler (0-12)
 * @returns {number}
 */
window.safeRound = function(num, decimals) {
    const n = Number(num);
    if (!Number.isFinite(n)) return NaN;
    const d = Math.max(0, Math.min(12, Math.floor(Number(decimals) || 0)));
    if (n === 0) return 0;

    const sign = n < 0 ? -1 : 1;
    const abs = Math.abs(n);

    // Eksponentiel skift undgår præcisionstab ved store tal som 1000.005
    const shifted = Number(abs + 'e+' + d);
    let result;
    if (Number.isFinite(shifted)) {
        result = sign * Number(Math.round(shifted) + 'e-' + d);
    } else {
        const factor = Math.pow(10, d);
        result = sign * (Math.round((abs + Number.EPSILON) * factor) / factor);
    }

    return Object.is(result, -0) ? 0 : result;
};

/**
 * Formaterer et tal til streng med valgt decimalseparator og afrunding.
 * Fjerner efterstillede nuller ("2.50" → "2,5") medmindre keepTrailingZeros=true.
 *
 * @param {number}  num               - Tal der skal formateres
 * @param {number}  decimals          - Antal decimaler
 * @param {string}  [decSep=',']      - Decimalseparator (',' eller '.')
 * @param {boolean} [keepTrailingZeros=false]
 * @returns {string}
 */
window.formatNumber = function(num, decimals, decSep, keepTrailingZeros) {
    const n = Number(num);
    if (num === null || num === undefined || typeof num === 'boolean' || !Number.isFinite(n)) {
        return "N/A";
    }
    const d = Math.max(0, Math.min(12, Math.floor(Number(decimals) ?? 2)));
    const sep = decSep || ',';
    const rounded = window.safeRound(n, d);
    if (!Number.isFinite(rounded)) return "N/A";

    let str = rounded.toFixed(d);
    if (!keepTrailingZeros) {
        str = str.replace(/(\.\d*[1-9])0+$/, '$1').replace(/\.0*$/, '');
    }
    if ( str === '-0') str = '0';
    if (sep === ',') {
        str = str.replace(/\./g, ',');
    }
    return str;
};
