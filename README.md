# SkoleCAS — Matematik-CAS direkte i Microsoft Word

<div align="center">

![Version](https://img.shields.io/badge/version-1.0.0-blue?style=for-the-badge)
![Platform](https://img.shields.io/badge/platform-Word%20Desktop%20%7C%20Word%20Online-0078d4?style=for-the-badge&logo=microsoft-word&logoColor=white)
![License](https://img.shields.io/badge/license-Fri%20til%20undervisning-green?style=for-the-badge)
![Sprog](https://img.shields.io/badge/sprog-Dansk-red?style=for-the-badge)
![Målgruppe](https://img.shields.io/badge/målgruppe-STX%20%7C%20HTX%20%7C%20HF%20%7C%20Folkeskole-orange?style=for-the-badge)

**Et gratis, open-source Computer Algebra System (CAS) der kører direkte i Microsoft Word.**  
Ingen installation. Ingen abonnement. Ingen ekstra programmer.

[🚀 Kom i gang](#-kom-i-gang-på-30-sekunder) · [📖 Funktioner](#-funktioner) · [🏗️ Arkitektur](#️-teknisk-arkitektur) · [❓ FAQ](#-faq--fejlsøgning)

</div>

---

## 💡 Hvad er SkoleCAS?

SkoleCAS er et **Microsoft Word-tilføjelsesprogram** (Add-in) designet til danske elever og lærere på folkeskole- og gymnasieniveau (STX, HTX, HF).

Marker en ligning eller et udtryk i dit Word-dokument — og med ét klik beregner, løser, differentierer eller integrerer SkoleCAS det for dig. Resultatet indsættes direkte som **rigtig matematiksrift** (ikke som billede eller plain tekst), præcis som det ser ud i en formelsamling eller lærebog.

---

## 🚀 Kom i gang på 30 sekunder

> [!IMPORTANT]
> SkoleCAS kræver Microsoft Word (Desktop eller Online). Det virker på Windows, Mac og i browseren.

1. **Download** [`manifest.xml`](manifest.xml) (højreklik → *Gem link som...*)
2. Åbn **Microsoft Word**
3. Gå til fanen **Indsæt** → klik **Tilføjelsesprogrammer**
4. Vælg **Flere tilføjelsesprogrammer** → fanen **Mine tilføjelsesprogrammer**
5. Klik **Overfør mit tilføjelsesprogram...** og vælg den downloadede `manifest.xml`
6. **SkoleCAS-panelet åbner i højre side** — du er klar! 🎉

---

## ✨ Funktioner

SkoleCAS har **seks faner**, der dækker de vigtigste matematikopgaver i gymnasiet og folkeskolen:

| Fane | Funktion | Nøglefunktioner |
|------|----------|-----------------|
| 🧮 **Beregner** | Inline-beregning i Word | Udregn, løs, differentier, integrer |
| 📐 **Trekant** | Trekantsberegner | Alle tilfælde inkl. SSA-tvetydighed |
| 📓 **Kladde** | Sandkasse-beregner | Eksperimenter uden at berøre dokumentet |
| 📈 **Graf** | GeoGebra graftegner | Interaktiv 2D-tegner, indsæt som billede |
| 📊 **Statistik** | Ugrupperet & grupperet | Stikprøvestatistik (s, s²), boksplot |
| 📉 **Regression** | Kurvetilpasning | Lineær, eksponentiel, potens, polynomium |

---

### 🧮 Beregner (Inline-tilstand)

Marker et matematisk udtryk i Word-dokumentet og tryk en af disse knapper:

| Knap | Handling | Eksempel |
|------|----------|---------|
| **Udregn (=)** | Beregn udtryk numerisk eller eksakt | `1/2 + 3/4` → `= 5/4` |
| **Løs (⇒)** | Løs ligning for den primære variabel | `2x + 5 = 15` → `⇒ x = 5` |
| **Diff. (f')** | Differentier mht. variablen | `x^3 + 2x` → `f'(x) = 3x² + 2` |
| **∫ Integrer** | Ubestemt integral | `2x` → `∫f(x)dx = x² + C` |

**Indstillinger:**

- ☑️ **Eksakt** — bevar brøker, π og kvadratrødder (slå fra for decimaltilnærmelse)
- 🔢 **Afrund** — vælg 0–5 decimaler (standard: 2)
- **,/.** — decimalseparator (komma er valgt som standard, se [afsnittet om inputformat](#-inputformat-og-decimaltegn))

**Symbolknapper:** Indsæt `π`, `√()`, `e` og `x²` direkte ved markørens position.

> [!NOTE]
> Beregnerresultater indsættes som **rigtig Word-matematiksrift (OOXML)** via en LaTeX → MathML → XSLT → OOXML pipeline — ikke som billeder eller plain tekst. Formlen ser ud som om den er skrevet med Words formeleditor.

---

### 📐 Trekantsberegner

Indtast præcis **3 kendte værdier** (sider og/eller vinkler). Vinkler angives i grader.

**Understøttede tilfælde:**

| Forkortelse | Beskrivelse | Understøttet |
|-------------|-------------|:------------:|
| SSS | Tre sider kendte | ✅ |
| SAS | To sider og den indesluttede vinkel | ✅ |
| ASA | To vinkler og en side | ✅ |
| AAS | To vinkler og en side (ikke indesluttet) | ✅ |
| SSA | To sider og modstående vinkel (tvetydigt) | ✅ med advarsel |
| AAA | Tre vinkler — ingen entydig trekant | ❌ (afvises) |

> [!WARNING]
> **SSA-tvetydighed («den dumme side»):** Når SSA giver *to* gyldige trekanter, indsætter SkoleCAS **begge løsninger** i dokumentet og viser en tydelig advarsel. Ingen løsning tabes lydigt.

**Validering:** Negative sider, vinkler uden for ]0°; 180°[ og vinkelsummer ≥ 180° giver præcise danske fejlbeskeder.

Resultater indsættes med fed overskrift og kursiv tekst i Word.

---

### 📓 Kladdehæftet

Et sandkasse-område til at eksperimentere frit, uden at noget skrives i Word-dokumentet med det samme. Skriv eller indsæt et udtryk, tryk **Beregn**, og overfør resultatet til dokumentet med ét klik på **Overfør til Word**.

---

### 📈 Graf (GeoGebra)

GeoGebra-graftegneren åbnes i en **popup-modal** der lægger sig oven på panelet:

- **📄 Hent fra Word** — henter markeret ligning og sender til GeoGebra automatisk
- **📈 Åbn graftegner** — åbner modal med fuld interaktiv GeoGebra-instans
- **↓ Indsæt i Word** — indsætter grafen som et skarpt PNG-billede i dokumentet

> [!NOTE]
> GeoGebra initialiseres *første gang* modalen åbnes og genbruges bagefter — du behøver ikke vente igen.  
> Lukker du modalen med **Esc** eller krydset, husker den din graf til næste gang i samme session.

---

### 📊 Statistik

**Ugrupperede data** — kopier tal direkte fra Excel. Fungerer med disse separatorer:

```
Komma:       12, 14, 15, 16, 12
Semikolon:   12; 14; 15; 16; 12
Linjeskift:  12
             14
             15
Mellemrum:   12 14 15 16 12
```

**Beregnede mål:**

| Mål | Symbol | Forklaring |
|-----|--------|------------|
| Antal | n | Datasættets størrelse |
| Minimum / Maksimum | Min / Max | Mindste og største observation |
| 1. og 3. kvartil | Q1 / Q3 | Kvartilinddeling |
| Median | Q2 | Midterste observation |
| Gennemsnit | x̄ | Aritmetisk middelværdi |
| Varians | **s²** | Stikprøvevarians — dividerer med *n*−1 |
| Spredning | **s** | Stikprøvespredning — √s² |

> [!IMPORTANT]
> **SkoleCAS bruger stikprøvestatistik (s og s²), som dividerer med n−1 — ikke n.**  
> Dette svarer nøjagtigt til, hvad der undervises i på STX, HTX og HF.  
> Brug **ikke** σ (sigma) som betegnelse — σ er populationsstatistik og dividerer med n.

Resultater indsættes som en **formateret Word-tabel** med boksplot-billede.

**Grupperede data** — angiv én klasse pr. linje i ét af disse formater:

```
# Format A: midtpunkt og frekvens
15 4
25 8
35 3

# Format B: intervalgrænser og frekvens (giver eksakt lineær interpolation)
10 20 4
20 30 8
30 40 3
```

Beregner gennemsnit, s, s², median og kvartiler med **lineær interpolation** inden for klasser (Format B).

---

### 📉 Regression

Kopier to kolonner (x og y) fra Excel og vælg regressionstype:

| Type | Modelformel | GeoGebra-kommando |
|------|-------------|-------------------|
| Lineær | f(x) = ax + b | `FitLine` |
| Eksponentiel | f(x) = b · eˢˣ | `FitGrowth` |
| Potens | f(x) = a · xᵇ | `FitPow` |
| Polynomium (2. grad) | f(x) = ax² + bx + c | `FitPoly(L₁, 2)` |

Indsætter **funktionsformel**, **R²-værdien** og et **plot** i Word.

> [!TIP]
> **Eksponentiel model:** GeoGebra bruger `b·eˢˣ`-formen. Konverter til `b·aˣ` med `a = eˢ`.

---

## ⌨️ Inputformat og decimaltegn

> [!IMPORTANT]
> **Du kan bruge komma som decimaltegn — det virker!**  
> `1,5` og `3,14` accepteres og konverteres automatisk til `1.5` og `3.14` inden beregning.

Koden bruger dette mønster (i `mathEngine.js` og `graphing.js`) til konverteringen:

```js
// Konverterer komma til punktum KUN når det står mellem to cifre
// (undgår at ødelægge funktionskald med kommaseparerede argumenter)
mathExpr = mathExpr.replace(/(\d),(\d)/g, '$1.$2');
```

| Du skriver | SkoleCAS forstår | Virker? |
|------------|-----------------|:-------:|
| `3,14` | `3.14` | ✅ |
| `1,5 + 2,5` | `1.5 + 2.5` | ✅ |
| `sin(30)` | `sin(30)` | ✅ |
| `x^2 + 3x` | `x^2 + 3x` | ✅ |
| `3.14` | `3.14` (punktum virker også) | ✅ |

**Understøttede symboler i input:**

| Symbol | Alternativ | Eksempel |
|--------|-----------|---------|
| `π` | skriv `pi` | `2*pi` eller `2π` |
| `√` | skriv `sqrt(` | `√4` eller `sqrt(4)` |
| `e` | skriv `e` | `e^2` |
| `^` | potensoperator | `x^3` |
| `*` | multiplikation | `2*x` |

---

## 🏗️ Teknisk arkitektur

```
SkoleCAS/
├── index.html          – UI: faner, knapper, modal-overlay, tema-skift
├── utils.js            – Delt modul: safeRound, normalizeInput, formatNumber
├── mathEngine.js       – Inline-beregning og kladdehæfte (Nerdamer CAS)
├── trigonometry.js     – Trekantsberegner med fuld SSA-håndtering
├── statistics.js       – Ugrupperet og grupperet statistik (simple-statistics)
├── regression.js       – Regressionsberegning via GeoGebra API
├── graphing.js         – GeoGebra modal-integration og graf-eksport
├── ommlConverter.js    – LaTeX → MathML (KaTeX) → OMML (XSLT) → Word OOXML
├── mml2omml.xsl        – Microsofts officielle MathML-til-OMML stylesheet
├── taskpane.css        – Styling med lyst og mørkt tema
└── manifest.xml        – Office Add-in manifest (version 1.0.0, locale da-DK)
```

<details>
<summary>📦 Afhængigheder (CDN-baserede biblioteker)</summary>

| Bibliotek | Version | Formål |
|-----------|---------|--------|
| [Office.js](https://appsforoffice.microsoft.com) | 1.1 | Word API — læs/skriv dokumentindhold |
| [Nerdamer-Prime](https://github.com/together-science/nerdamer-prime) | 1.4.0 | CAS-motor (algebra, calculus, ligningsløsning) — aktivt vedligeholdt fork af nerdamer |
| [KaTeX](https://katex.org) | 0.16.8 | LaTeX → MathML konvertering |
| [GeoGebra](https://geogebra.org) | latest | Graftegner og regressionsberegning |
| [Chart.js](https://chartjs.org) | 4 | Boksplot-tegning |
| [chartjs-chart-boxplot](https://github.com/sgratzl/chartjs-chart-boxplot) | 4 | Boksplot-plugin til Chart.js |
| [simple-statistics](https://simplestatistics.org) | 7.8.2 | Statistiske beregninger (sampleVariance, quantile osv.) |

</details>

<details>
<summary>🔄 Matematisk pipeline for Word-indsættelse</summary>

Hvert beregnet resultat gennemgår denne pipeline inden det indsættes i Word:

```
Brugerens tekst (f.eks. "1/2 + 3/4")
    │
    ▼
normalizeInput()          ← utils.js: π→pi, √→sqrt
    │
    ▼
komma-til-punktum regex   ← (\d),(\d) → $1.$2
    │
    ▼
Nerdamer CAS              ← Algebra, Calculus, Solve
    │
    ▼
.toTeX()                  ← LaTeX-streng
    │
    ▼
KaTeX renderToString()    ← MathML-streng
    │
    ▼
XSLTProcessor             ← OMML via mml2omml.xsl
    │
    ▼
range.insertOoxml()       ← Word API (indsæt som rigtig formel)
```

Hvis OOXML-pipelinen fejler, falder systemet elegant tilbage til at indsætte Unicode-tekst (f.eks. `≈`, `⇒`, `∫`, `π`).

</details>

<details>
<summary>🧩 utils.js — det delte hjælpemodul</summary>

`utils.js` indlæses **som det første script** og eksporterer tre globale funktioner til alle øvrige moduler:

```js
// Normaliserer matematisk input (π, Π → pi  og  √ → sqrt)
window.normalizeInput(expr: string): string

// Sikker afrunding der undgår IEEE 754 floating-point fejl
// Eksempel: safeRound(1.005, 2) → 1.01  (ikke 1.00 som standard toFixed giver)
window.safeRound(num: number, decimals: number): number

// Formaterer tal til streng med korrekt decimalseparator og afrunding
// Fjerner trailing nuls (2.50 → 2,5) medmindre keepTrailingZeros=true
window.formatNumber(num: number, decimals: number, decSep: string, keepTrailingZeros?: boolean): string
```

</details>

---

## 🖥️ Lokal udvikling

Åbn projektmappen i terminalen og start en lokal HTTPS-server:

```bash
# Med Python (følger med på de fleste systemer):
python -m http.server 8080

# Eller med Node.js (npx):
npx serve .
```

Åbn `http://localhost:8080` i browseren for at se UI'et.

> [!WARNING]
> Word API'et kræver **HTTPS** i produktionsmiljøet. Til lokal test bruges Word Desktop med sideload via `manifest.xml`.  
> Word Online kræver at filen hostes på en offentlig HTTPS-URL (f.eks. GitHub Pages).

### Hosting på GitHub Pages

`manifest.xml` peger på GitHub Pages. Ret `<SourceLocation>` og `<SupportUrl>` til din URL:

```xml
<SupportUrl DefaultValue="https://DIT-BRUGERNAVN.github.io/SkoleCAS/index.html" />
<SourceLocation DefaultValue="https://DIT-BRUGERNAVN.github.io/SkoleCAS/index.html" />
```

---

## ⚠️ Kendte begrænsninger

| Begrænsning | Beskrivelse | Workaround |
|-------------|-------------|-----------|
| **nerdamer-prime v1.4.0** | Aktivt vedligeholdt fork af det arkiverede nerdamer. Har begrænsninger med komplekse tal og visse specialintegraler | Tilstrækkeligt for gymnasieniveau |
| **Variabelnavne** | `e`, `pi`, `sin`, `cos`, `tan`, `log` er reserverede nøgleord | Brug `x` som primær variabel |
| **Eksponentiel regression** | GeoGebra returnerer `b·eˢˣ`, ikke `b·aˣ` | Beregn `a = eˢ` manuelt |
| **Netadgang** | Kræver internetforbindelse til CDN-biblioteker og GeoGebra | Fungerer ikke rent offline |
| **Word API HTTPS** | Word Online kræver HTTPS-hosting | Brug GitHub Pages eller tilsvarende |

---

## ❓ FAQ — Fejlsøgning

<details>
<summary><strong>Panelet åbner ikke eller viser en hvid skærm</strong></summary>

1. Kontrollér at din `manifest.xml` peger på en tilgængelig HTTPS-URL
2. Ryd browser-cachen i Word Online (Ctrl+Shift+R)
3. Prøv at fjerne og geninstallere tilføjelsesprogrammet
4. Tjek at du har en aktiv internetforbindelse (CDN-biblioteker indlæses ved opstart)

</details>

<details>
<summary><strong>«Matematisk fejl» eller «Kunne ikke beregne»</strong></summary>

- Brug `x` som variabelnavn (ikke `t`, `k`, `n` osv.) i inline-beregneren
- `e`, `pi`, `sin`, `cos`, `tan`, `log` er reserverede — brug dem ikke som variabelnavne
- Tjek at udtrykket er syntaktisk korrekt: `2*x + 5 = 15` (ikke `2x+5=15` uden mellemrum kan fejle i kanttilfælde)
- Potenser skrives med `^`: `x^2` ikke `x²`

</details>

<details>
<summary><strong>Komma i input virker ikke / forkerte decimaler</strong></summary>

SkoleCAS konverterer **automatisk** `1,5` til `1.5` — du behøver ikke selv tænke over det.

Konverteringen sker med `replace(/(\d),(\d)/g, '$1.$2')`, som kun erstatter kommaer der står **imellem to cifre**. Det betyder:
- ✅ `3,14` → `3.14`
- ✅ `1,5 + 2,5` → `1.5 + 2.5`
- ✅ `nerdamer('solve', x+1=0)` — kommaet i funktionskald berøres **ikke**

Hvis du oplever problemer, tjek om du har valgt den rigtige decimalseparator i indstillingerne (skruetrækker-ikon i beregner-fanen).

</details>

<details>
<summary><strong>GeoGebra starter ikke / graf-modalen er tom</strong></summary>

1. GeoGebra hentes fra `cdn.geogebra.org` — tjek din internetforbindelse
2. Vent op til 15 sekunder første gang (GeoGebra er et stort bibliotek)
3. Hvis der vises «GeoGebra kunne ikke starte» — genindlæs panelet
4. GeoGebra-modalen åbner som popup oven på panelet — tjek at pop-ups ikke er blokeret

</details>

<details>
<summary><strong>Statistik: mine tal matcher ikke TI-Nspire / Excel</strong></summary>

SkoleCAS bruger **stikprøvestatistik** (dividerer med *n*−1). Mange lommeregnere og Excel kan skiftes mellem stikprøve- og populationsstatistik:

| Funktion | Dividerer med | SkoleCAS? |
|----------|--------------|:---------:|
| s (stikprøvespredning) | n−1 | ✅ Bruges |
| σ (populationsspredning) | n | ❌ Bruges ikke |

I Excel: brug `STDEV` / `VAR` (svarer til s og s²), **ikke** `STDEVP` / `VARP`.  
I TI-Nspire: vælg `Sx` og `Sx²` i statistikmenuen, ikke `σx`.

</details>

<details>
<summary><strong>Trekantsberegneren giver to løsninger</strong></summary>

Det er **korrekt opførsel** — du har givet et SSA-tilfælde (to sider og en modstående vinkel), som matematisk kan give to gyldige trekanter. Begge løsninger er gyldige og indsættes begge i dokumentet med en tydelig advarsel. Aflæs begge og vurdér hvilken der passer til din opgave.

</details>

<details>
<summary><strong>Regression: «GeoGebra starter...» fejlbesked</strong></summary>

Regression bruger GeoGebra som beregningsmotor. Første gang du trykker «Beregn Regression», åbnes GeoGebra-modalen i baggrunden for at initialisere. Vent ca. 5–10 sekunder og tryk på knappen igen.

</details>

---

## 🤝 Bidrag

Bidrag, fejlmeldinger og forslag er velkomne!

1. **Fork** dette repository
2. Opret en **ny branch**: `git checkout -b feature/mit-bidrag`
3. Foretag dine ændringer og **commit**: `git commit -m "Tilføj: beskrivelse af ændring"`
4. **Push** branchen: `git push origin feature/mit-bidrag`
5. Åbn en **Pull Request** med en beskrivelse af hvad du har ændret og hvorfor

### Retningslinjer for kode

- Skriv kommentarer på **dansk** (projektet er til danske skoler)
- Brug `window.safeRound()` fra `utils.js` frem for `toFixed()` direkte — undgår IEEE 754-fejl
- Brug `window.normalizeInput()` fra `utils.js` til al input-normalisering — én kilde til sandhed
- Brug `window.formatNumber()` fra `utils.js` til tal-formatering med decimaltegn
- Test i **både Word Desktop og Word Online** inden du sender PR

### Hvad vi gerne vil have hjælp til

- [ ] Understøttelse af flere variabler (f.eks. `f(x, y)`)
- [ ] Understøttelse af vektorer og matricer
- [ ] Offline-tilstand (bundlede biblioteker)
- [ ] Formelhistorik i kladdehæftet
- [ ] Eksport af statistikoversigt til Excel

---

## 📜 Licens

SkoleCAS er frit at bruge, kopiere og modificere til **undervisningsformål**.  
Kommerciel brug kræver tilladelse fra forfatteren.

---

<div align="center">

Bygget med ❤️ til den digitale matematikundervisning i Danmark.

*SkoleCAS v1.0.0 · Office Add-in ID: `3a4cb1b0-1234-4a4a-9b9b-c2e505a76c02` · Hosted på [GitHub Pages](https://gancia.github.io/SkoleCAS/index.html)*

</div>
