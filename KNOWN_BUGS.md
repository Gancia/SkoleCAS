# Status på kendte fejl og tekniske løsninger i SkoleCAS

Dette dokument beskriver de tekniske udfordringer i Microsoft Word Add-in-miljøet (særligt Word Online) og hvordan de er løst i den aktuelle arkitektur.

## 1. Låste ligningsfelter og `a8pId`-fejl i Word Online (Løst)

### Baggrund
I Word Online deler et nyt afsnit efter `Enter` midlertidigt `paraId` med det forrige afsnit. For at tvinge Word Online til at oprette et unikt afsnit indsættes midlertidigt en Content Control (`<w:sdt>`), som udskiftes med et `<m:oMath>`-element via `paragraph.insertOoxml(newOoxml, "Replace")`.

Tidligere opstod to problemer i Word Online:
1. `Cannot read properties of null (reading 'a8pId')` ved kald til `targetRange.search(...)` eller genbrug af gamle Office.js `Range`-referencer efter `insertOoxml`.
2. Låste ligningsfelter, hvor markøren ikke kunne forlade `<m:oMath>`-feltet eller slette boksen, fordi Word Online beholdt en usynlig Content Control-wrapper omkring ligningen.

### Løsning i `mathEngine.js` (`WordBridge.insertMathInWord`)
- Alle kald til `.search()` på midlertidige `Range`-objekter er fjernet.
- Efter `<m:oMath>` indsættes altid et almindeligt tekst-run (`<w:r><w:t xml:space="preserve"> </w:t></w:r>`), så brugerens markør kan bevæge sig frit ud til højre for ligningen.
- Efter `paragraph.insertOoxml(newOoxml, "Replace")` hentes en frisk markering (`context.document.getSelection()`), og eventuelle tilbageværende `MATH_TARGET_`-Content Controls eller `parentContentControlOrNullObject` låses op og pakkes ud med `.delete(true)` (`keepContent = true`), hvorefter markøren flyttes til `"End"`.

---

## 2. Tovejs OMML-oversættelse uden XSLT (Løst)

### Baggrund
Browserens indbyggede `XSLTProcessor` og filen `MML2OMML.XSL` er ustabile på tværs af moderne WebViews i Office Online og er under udfasning i Chromium.

### Løsning i `ommlConverter.js`
- **Skrivning (`window.latexToOmmlString`):** En ren JavaScript LaTeX-til-OMML-kompiler bygger direkte `<m:oMath>`, `<m:f>` (brøker), `<m:rad>` (rødder), `<m:sSup>` / `<m:sSub>` / `<m:sSubSup>` (potenser og indeks) samt `<m:d>` (parenteser) og samler danske decimaltal (`3,14`) i ét `<m:r><m:t>3,14</m:t></m:r>`.
- **Læsning (`window.ommlToExpression`):** En rekursiv DOM-parser læser eksisterende `<m:oMath>`-ligninger tilbage fra Word til normaliseret matematik-syntaks, så elever kan markere en eksisterende Word-ligning og klikke `Løs (⇒)`, `Differentier (f')` eller `Integrer (∫)` direkte.

---

## 3. Strukturel LaTeX-forhåndsvisning uden forhåndsevaluering (Løst)

### Baggrund
Hvis venstresiden af et udtryk (f.eks. `4 + 6`, `1/2 + 3/4` eller `√(25) + 5`) sendes gennem Nerdamer for at generere LaTeX, reducerer Nerdamer udtrykket til `10` eller `5/4`, så selve mellemregningen forsvinder, og der blot står `10 = 10`.

### Løsning i `mathEngine.js` (`CAS.formatSegmentToLatex` & `CAS.toPreviewLatex`)
- Venstresiden af udtrykket og live-forhåndsvisningen formateres strukturelt til LaTeX uden at evaluere talværdierne. Dermed vises og indsættes `\sqrt{25} + 5 = 10` og `\frac{1}{2} + \frac{3}{4} = \frac{5}{4}` korrekt.

---

## 4. Afrunding, brøk-parsing og trekantsberegner (Løst)

- **Negativ afrunding (`roundNumbersInString`):** Regexen matcher kun positive decimaltal (`/\d+\.\d+/g`), så et lille negativt led som `-0.001*x` afrundes til `-0*x` i stedet for at sluge minustegnet.
- **Brøker med potenser og funktioner (`formatFractions`):** Tæller og nævner scannes med balanceret parentes- og tuborgklamme-genkendelse, så `(x+1)^2/3`, `\sin(x)/2`, `1/(x+1)^2` og `1/\sin(x)` omsættes korrekt til `\frac{...}{...}`.
- **Tvetydigt SSA-tilfælde i `trigonometry.js`:** Når en trekant har to gyldige løsninger, vises begge løsninger med hver sin målestoksrette SVG-skitse i panelet, og ved automatisk indsættelse i Word afventes hver indsættelse sekventielt (`await`), så rækkefølgen altid er Løsning 1 efterfulgt af Løsning 2.
