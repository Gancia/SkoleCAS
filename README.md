# SkoleCAS — Matematik-CAS direkte i Microsoft Word

<div align="center">

![Version](https://img.shields.io/badge/version-2.0.0-blue?style=for-the-badge)
![Platform](https://img.shields.io/badge/platform-Word%20Desktop%20%7C%20Word%20Online-0078d4?style=for-the-badge&logo=microsoft-word&logoColor=white)
![License](https://img.shields.io/badge/license-Fri%20til%20undervisning-green?style=for-the-badge)
![Sprog](https://img.shields.io/badge/sprog-Dansk-red?style=for-the-badge)
![Målgruppe](https://img.shields.io/badge/målgruppe-Folkeskole%20%7C%20STX%20%7C%20HTX%20%7C%20HF-orange?style=for-the-badge)

**Et gratis, open-source Computer Algebra System (CAS) der kører direkte i Microsoft Word.**  
Ingen lokal installation. Ingen licensomkostninger. Bygget til dansk matematikundervisning.

[Kom i gang](#kom-i-gang) · [Arbejdsflow](#arbejdsflow-i-word-og-panelet) · [Funktioner](#funktioner) · [Teknisk arkitektur](#teknisk-arkitektur) · [FAQ](#faq--fejlsøgning)

</div>

---

## Hvad er SkoleCAS?

SkoleCAS er et **Microsoft Word-tilføjelsesprogram** (Office Add-in) målrettet elever og lærere i grundskolen og på gymnasiale uddannelser (STX, HTX, HF). Formålet er at levere en hurtig, stabil og gratis pendant til de mest brugte hverdagsværktøjer fra WordMat direkte i både Word Desktop (Windows/Mac) og Word Online.

Eleverne kan arbejde direkte i deres Word-dokument eller bruge panelet i højre side med øjeblikkelig matematisk forhåndsvisning. Alle udregninger indsættes som **ægte Office Math-formler (`<m:oMath>`)**, der kan redigeres videre i Words indbyggede ligningseditor (`Alt + =`).

---

## Kom i gang

> [!IMPORTANT]
> SkoleCAS fungerer i Microsoft Word på Windows, Mac og i browseren (Word Online).

1. **Download** [`manifest.xml`](manifest.xml) (højreklik og vælg *Gem link som...*)
2. Åbn **Microsoft Word**
3. Gå til fanen **Start** (eller **Hjem**) og klik på **Tilføjelsesprogrammer**
4. Vælg **Flere tilføjelsesprogrammer** og skift til fanen **Mine tilføjelsesprogrammer**
5. Klik på **Overfør mit tilføjelsesprogram...** og vælg den downloadede `manifest.xml`
6. **SkoleCAS-panelet åbner i højre side** og er klar til brug.

---

## Arbejdsflow i Word og panelet

SkoleCAS understøtter et **tovejs arbejdsflow**, så eleverne kan vælge den arbejdsform, der passer dem bedst:

1. **Direkte i Word-dokumentet (som i WordMat):**
   - Skriv et regnestykke eller en ligning på en linje i Word (enten som almindelig tekst eller inde i et Word-ligningsfelt med `Alt + =`), eller markér en eksisterende ligning, og klik på **Udregn (=)** eller **Løs (⇒)**.
   - SkoleCAS aflæser automatisk linjen eller markeringen (inklusive ægte Word-brøker, potenser og kvadratrødder), viser mellemregningen i panelet og indsætter resultatet som en rigtig Word-ligning direkte i dokumentet.
2. **I panelets matematikfelt (med live forhåndsvisning):**
   - Skriv udtrykket i feltet i panelet, hvor det formateres løbende med rigtige brøkstreger, potenser og rodtegn, før du trykker `Enter`.
   - Resultatet vises i resultatkortet, gemmes i sessionshistorikken (de seneste 8 udregninger) og kan indsættes i Word automatisk eller med et enkelt klik.

---

## Funktioner

SkoleCAS er bygget op omkring to fokuserede kernemoduler samt en direkte bro til GeoGebra Online:

| Modul | Formål | Nøglefunktioner |
|-------|--------|-----------------|
| **Beregner & Ligningsløser** | CAS-beregning og algebra | Udregn (`=`), løs ligning (`⇒`), differentier (`f'`), integrer (`∫`), live forhåndsvisning og historik |
| **Trekantsberegner** | Trigonometri i vilkårlige trekanter | SSS, SAS, ASA, AAS og SSA (inkl. dobbeltløsning), areal, omkreds og målfast SVG-skitse |
| **GeoGebra Online-bro** | Graftegning og geometri | Automatisk konvertering af dansk syntaks til GeoGebra, kopiering til udklipsholder og kommandoguide |

---

### 1. Beregner & Ligningsløser

Skriv et udtryk i Word eller i panelet og vælg en af de fire handlinger:

| Knap | Handling | Eksempel |
|------|----------|----------|
| **Udregn (`=`)** | Udregner eller reducerer udtryk numerisk eller eksakt | `1/2 + 3/4` $\rightarrow$ `\frac{1}{2} + \frac{3}{4} = \frac{5}{4}` |
| **Løs (`⇒`)** | Løser ligning for den automatisk fundne variabel | `2x + 5 = 15` $\rightarrow$ `2x + 5 = 15 \Rightarrow x = 5` |
| **Differentier (`f'`)** | Differentierer udtrykket mht. variablen | `x^3 + 2x` $\rightarrow$ `x^{3} + 2x \Rightarrow f'(x) = 3x^{2} + 2` |
| **Integrer (`∫`)** | Finder det ubestemte integral inkl. konstantled | `2x` $\rightarrow$ `2x \Rightarrow \int f(x)\,dx = x^{2} + C` |

**Indstillinger for beregning:**
- **Eksakt resultat (bevar brøker, π og rødder):** Bevarer eksakte brøker, $\pi$ og kvadratrødder. Slås det fra, beregnes decimaltilnærmelse ($\approx$).
- **Vis brøker stablet (lodret):** Vælger mellem opstillede brøker ($\frac{a}{b}$) og linjebrøker ($a/b$).
- **Indsæt automatisk i Word ved beregning:** Bestemmer om beregninger straks skrives ind ved markøren i Word, eller kun vises i resultatkortet indtil man klikker **Indsæt i Word**.
- **Decimaler & decimalseparator:** Vælg 0–5 decimaler (standard: 2) samt dansk komma (`,`) eller punktum (`.`).

---

### 2. Bro til GeoGebra Online

I stedet for at tvinge en tung graftegner ind i det smalle Word-panel har SkoleCAS en dedikeret bro til **GeoGebra Online**:

- Klik på **Åbn i GeoGebra Online** for automatisk at omskrive det aktuelle udtryk fra dansk notation (`f(x) = 2,5x² - 3π`) til gyldig GeoGebra-syntaks (`f(x) = 2.5x^2 - 3*pi`).
- Udtrykket kopieres automatisk til udklipsholderen, så eleven blot skal trykke `Ctrl + V` i GeoGebra.
- Panelet indeholder desuden en oversigt over de vigtigste danske GeoGebra-kommandoer (`Skæring`, `Rod`, `Ekstremum`, `Tangent`, `Integral`) samt vejledning i billed-eksport til Word.

---

### 3. Trekantsberegner

Indtast præcis **3 kendte værdier** (sider $a, b, c$ og/eller vinkler $A, B, C$ i grader). SkoleCAS beregner de resterende sider og vinkler ved hjælp af cosinus- og sinusrelationerne samt trekantens **areal ($T$)** og **omkreds ($O$)**.

| Tilfælde | Beskrivelse | Understøttet |
|----------|-------------|:------------:|
| **SSS** | Tre kendte sider | Ja |
| **SAS** | To sider og den indesluttede vinkel | Ja |
| **ASA / AAS** | To vinkler og én side | Ja |
| **SSA** | To sider og en modstående vinkel (tvetydigt tilfælde) | Ja (viser begge løsninger) |
| **AAA** | Tre vinkler uden kendt side | Afvises med forklaring |

> [!WARNING]
> **SSA-tvetydighed:** Når to sider og en modstående vinkel matematisk giver to gyldige trekanter ($B_1$ spids og $B_2$ stump), viser SkoleCAS en tydelig advarsel, tegner en **målfast SVG-skitse** for hver af de to trekanter og lader eleven indsætte enten én eller begge løsninger i Word.

---

## Inputformat og dansk komma

SkoleCAS accepterer både dansk komma (`,`) og punktum (`.`) som decimaltegn samt almindelige matematiske genvejstegn:

| Du skriver | SkoleCAS forstår | Understøttet |
|------------|------------------|:------------:|
| `3,14` | `3.14` (bevarer `3,14` i visning og Word) | Ja |
| `2x^2 + 3x` eller `2x² + 3x` | $2x^2 + 3x$ (implicit multiplikation og potens) | Ja |
| `sqrt(x + 1)` eller `√(x + 1)` | $\sqrt{x + 1}$ | Ja |
| `2*pi` eller `2π` | $2\pi$ | Ja |
| `sin(x)`, `cos(x)`, `tan(x)`, `ln(x)` | Standardfunktioner | Ja |

---

## Teknisk arkitektur

```text
SkoleCAS/
├── index.html          – Brugergrænseflade med 2 faner, symbolbjælke, resultatkort og GeoGebra-hjælp
├── taskpane.css        – Komponentbaseret styling med lyst og mørkt tema til Word Taskpane (300–460px)
├── utils.js            – Input-normalisering, dansk kommahåndtering og eksponentiel safeRound()
├── ommlConverter.js    – Tovejs Pure-JS OMML-oversætter (LaTeX <-> Word <m:oMath>, 0 KB XSLT)
├── mathEngine.js       – Lagdelt kerne: CAS-adapter, WordBridge, GeoGebraBridge og UI/historik
├── trigonometry.js     – Trekantsberegner med SSA-håndtering, areal/omkreds og målfast SVG-tegner
├── main.js             – Initialisering af Office.js, faneskift, tema og hændelsesstyring
└── manifest.xml        – Office Add-in manifest (da-DK)
```

### De 4 arkitekturlag i `mathEngine.js` & `ommlConverter.js`

1. **`window.CAS` (`mathEngine.js`):**
   - Isoleret CAS-adapter omkring Nerdamer-Prime. Samler al udtryksanalyse, variabeldetektion, eksakt/numerisk evaluering og strukturel LaTeX-formatering (`formatSegmentToLatex`) ét sted, så nye matematiske moduler nemt kan kobles på.
2. **Tovejs Pure-JS OMML-bro (`ommlConverter.js`):**
   - **Ind i Word (`window.latexToOmmlString`):** En direkte *recursive-descent* LaTeX-til-OMML compiler skrevet i ren JavaScript uden brug af `XSLTProcessor` eller `mml2omml.xsl`. Genererer `<m:f>`, `<m:rad>`, `<m:sSup>`, `<m:sSub>`, `<m:sSubSup>` og skalerende `<m:d>`-parenteser, og holder danske decimaltal (`3,14`) samlet i ét `<m:t>`-element.
   - **Ud af Word (`window.ommlToExpression`):** Parser `<m:oMath>`-strukturer direkte fra Words OOXML (`selection.getOoxml()`), så formler skrevet med Words egen formeleditor (`Alt + =`) kan læses og genberegnes uden tab af brøkstreger eller potenser.
3. **`window.WordBridge` (`mathEngine.js`):**
   - Styrer al læsning og skrivning til Word-dokumentet. Løser låsefejlen i Word Online ved at placere `<m:oMath>` inde i `<w:sdtContent>` under `paragraph.insertOoxml(..., "Replace")` og derefter pakke Content Control-rammen rent ud via `item.delete(true)`. Undgår desuden indlejrede `<m:oMath>`-tags, hvis markøren allerede står inde i et ligningsfelt.
4. **`window.GeoGebraBridge` (`mathEngine.js`):**
   - Oversætter dansk matematik-notation til GeoGebra-syntaks, kopierer til udklipsholderen og åbner GeoGebra Online i en ny fane.

### Eksterne biblioteker (CDN)

| Bibliotek | Version | Formål |
|-----------|---------|--------|
| [Office.js](https://appsforoffice.microsoft.com) | 1.1 | Microsoft Word JavaScript API |
| [Nerdamer-Prime](https://github.com/together-science/nerdamer-prime) | 1.4.0 | Symbolsk CAS-motor (algebra, ligningsløsning, differential- og integralregning) |
| [KaTeX](https://katex.org) | 0.16.8 | Hurtig LaTeX-til-HTML matematikvisning i panelet |

---

## Lokal udvikling

Åbn projektmappen i terminalen og start en lokal webserver:

```bash
# Med Python:
python -m http.server 8080

# Eller med Node.js:
npx serve .
```

Åbn `http://localhost:8080` i browseren for at teste panelets brugergrænseflade og CAS-beregninger direkte i browser-tilstand.

> [!NOTE]
> Ved brug inde i Microsoft Word Online kræves det, at add-in'et hostes via **HTTPS** (f.eks. på GitHub Pages). Til lokal test i Word Desktop kan `manifest.xml` sideloades direkte.

### Hosting på GitHub Pages

Ret `<SourceLocation>` og `<SupportUrl>` i `manifest.xml` til din egen GitHub Pages-adresse:

```xml
<SupportUrl DefaultValue="https://DIT-BRUGERNAVN.github.io/SkoleCAS/index.html" />
<SourceLocation DefaultValue="https://DIT-BRUGERNAVN.github.io/SkoleCAS/index.html" />
```

---

## FAQ — Fejlsøgning

<details>
<summary><strong>Panelet åbner ikke eller viser en tom skærm</strong></summary>

1. Kontrollér at `manifest.xml` peger på en gyldig HTTPS-adresse.
2. Genindlæs panelet eller ryd browser-cachen i Word Online (`Ctrl + Shift + R`).
3. Tjek at der er adgang til internettet, så Office.js, Nerdamer-Prime og KaTeX kan indlæses.

</details>

<details>
<summary><strong>Hvordan regner jeg videre på en ligning i Word?</strong></summary>

Du kan enten markere ligningen i Word og klikke direkte på en af handlingsknapperne (**Beregn**, **Løs ligning**, **Differentier**, **Integrer**) eller klikke på **Hent fra Word** for at trække ligningen ind i panelets tekstfelt, rette i den og genberegne.

</details>

<details>
<summary><strong>Trekantsberegneren viser to løsninger</strong></summary>

Når du indtaster to sider og en vinkel, der ikke ligger imellem dem (SSA), kan der matematisk eksistere to forskellige gyldige trekanter. SkoleCAS viser begge løsninger med hver sin målfaste skitse, så du kan sammenligne dem med din opgaveformulering.

</details>

---

## Bidrag og kodeprincipper

Bidrag, fejlmeldinger og forbedringsforslag er velkomne via Issues og Pull Requests.

- Skriv kommentarer og brugervendte tekster på **dansk**.
- Hold brugergrænseflade, kode og dokumentation **fri for emojis** — brug i stedet rene SVG-ikoner i brugergrænsefladen.
- Brug altid `window.safeRound()`, `window.normalizeInput()`, `window.parseDanishNumber()` og `window.formatNumber()` fra `utils.js` for konsistent tal- og kommahåndtering.
- Hold CAS-kald samlet i `window.CAS` og Word-specifikke API-kald samlet i `window.WordBridge`.

---

## Licens

SkoleCAS kan frit anvendes, kopieres og videreudvikles til **undervisningsformål**.  
Kommerciel brug kræver forudgående aftale.
