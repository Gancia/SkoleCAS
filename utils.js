/**
 * SkoleCAS – Delte hjælpefunktioner
 * Indlæses som det FØRSTE script (defer, men øverst i listen) for at sikre
 * at alle hjælpefunktioner er tilgængelige i de øvrige moduler.
 */

/**
 * Normaliserer et matematisk udtryk til Nerdamer/GeoGebra-format.
 * Erstatter π, Π og √ med ASCII-ækvivalenter, så alle moduler bruger
 * samme logik ét sted frem for at duplikere replace()-kæder.
 * @param {string} expr - Rå matematisk streng fra brugeren
 * @returns {string}
 */
window.normalizeInput = function(expr) {
    return expr
        .replace(/[πΠ]/g, 'pi')
        .replace(/√\(/g, 'sqrt(')
        .replace(/√([a-zA-Z0-9.]+)/g, 'sqrt($1)');
};

/**
 * Sikker afrunding som undgår klassiske IEEE 754 floating-point fejl.
 *
 * Eksempel på fejl i standard toFixed:
 *   (1.005).toFixed(2) → "1.00"  ← FORKERT (fordi 1.005 ikke kan repræsenteres
 *                                    præcist i binær flydende-komma)
 *
 * Denne funktion løser det ved at tilføje Number.EPSILON som en lille korrektion
 * inden afrundingen, så kanttilfælde runder op som forventet:
 *   safeRound(1.005, 2) → 1.01  ← KORREKT
 *
 * @param {number} num      - Tal der skal afrundes
 * @param {number} decimals - Antal ønskede decimaler (0-15)
 * @returns {number}
 */
window.safeRound = function(num, decimals) {
    if (!isFinite(num)) return num;
    const factor = Math.pow(10, decimals);
    // C6: For at undgå IEEE 754-fejl skal korrektionen ALTID runde væk fra nul.
    // Math.sign(num) sikrer at vi adderer +epsilon for positive og -epsilon for negative tal.
    // Uden dette giver safeRound(-1.005, 2) = -1.00 i stedet for det korrekte -1.01.
    return Math.round((num + Math.sign(num) * Number.EPSILON) * factor) / factor;
};

/**
 * Formaterer et tal til en streng med korrekt decimalseparator og afrunding.
 * Fjerner trailing nuls (f.eks. "2.50" → "2,5") medmindre keepTrailingZeros=true.
 *
 * @param {number}  num               - Tal der skal formateres
 * @param {number}  decimals          - Antal decimaler
 * @param {string}  decSep            - Decimalseparator: ',' eller '.'
 * @param {boolean} [keepTrailingZeros=false] - Behold trailing nuls (f.eks. for tabeller)
 * @returns {string}
 */
window.formatNumber = function(num, decimals, decSep, keepTrailingZeros) {
    const rounded = window.safeRound(num, decimals);
    let str = rounded.toFixed(decimals);
    if (!keepTrailingZeros) {
        // Fjern trailing nuls, men behold mindst ét ikke-nul decimal
        str = str.replace(/(\.\d*[1-9])0+$/, '$1').replace(/\.0*$/, '');
    }
    if (decSep === ',') {
        // O7: Brug regex (med 'g'-flag) ikke en string-replace så ALLE punktummer erstattes,
        // ikke kun det første. Stræng-replace ('.' -> ',') håndterer kun én forekomst.
        str = str.replace(/\./g, ',');
    }
    return str;
};
