# Known Bugs and Code Quality Issues

This document tracks known bugs, edge cases, performance bottlenecks, and code quality improvements identified during a comprehensive codebase review. 

## Table of Contents
- [main.js](#mainjs)
- [utils.js](#utilsjs)
- [index.html](#indexhtml)
- [mathEngine.js](#mathenginejs)
- [graphing.js](#graphingjs)
- [statistics.js](#statisticsjs)
- [regression.js](#regressionjs)
- [taskpane.css](#taskpanecss)
- [ommlConverter.js](#ommlconverterjs)
- [trigonometry.js](#trigonometryjs)

---

## `main.js`

### 🐞 Bugs & Edge Cases
* **Lingering Error Messages:** In `insertSymbol` and `insertToWord`, `errorBox.innerText` is populated on error but never cleared on subsequent successful executions.
* **Footer Target Section:** `ensureFooter` inserts the footer into `context.document.sections.getFirst()` instead of the active section (`context.document.getSelection().sections.getFirst()`).
* **Race Condition with Office Initialization:** UI interaction events are bound on `DOMContentLoaded` before `Office.onReady` resolves. Clicking buttons early throws errors.
* **Missing Type Coercion:** `insertToWord` passes `window.currentResult` directly to `range.insertText()`. If it's a number/object, it crashes.
* **Silent Failures in Footer:** Errors in `ensureFooter` are caught and logged, leaving the user uninformed if permissions or protection block insertion.

### ⚡ Performance Issues
* **Redundant DOM Queries:** The same `document.getElementById()` is repeatedly called inside the `DOMContentLoaded` listener instead of caching.
* **Repeated Querying in `switchTab`:** `document.querySelectorAll` is executed on every tab switch.
* **Unnecessary Final Sync:** `insertToWord` calls `await context.sync()` redundantly after `ensureFooter` already syncs.

### 🧹 Code Quality
* **Global Namespace Pollution:** Extensive use of `window.currentResult`, `window.switchTab`, etc.
* **Outdated Declarations:** Use of `var` in loops instead of block-scoped `let`/`const`.
* **Inefficient Event Binding:** Use of `element.onclick` over `addEventListener`.
* **Repetitive DOM Bindings:** Duplicate blocks for binding `calculateInline` to various solve buttons.

---

## `utils.js`

### 🐞 Bugs & Edge Cases
* **Type Safety in `normalizeInput`:** Calling `.replace()` on `null`, `undefined`, or numbers throws a `TypeError`.
* **Regex Limits in `normalizeInput`:** The `/√([a-zA-Z0-9.]+)/g` regex fails on negative numbers or spaces (e.g., `√-5`, `√ 4`). The `/π/g` regex misses alternative unicode variations like `𝜋`.
* **Critical Bug in `safeRound`:** Epsilon scaling (`num + Math.sign(num) * Number.EPSILON`) fails for larger numbers (e.g., `1000.005`) due to floating-point absorption.
* **Type Safety in `formatNumber`:** `isFinite("abc")` returns `false`, causing `.toFixed()` to crash on strings. `isFinite(null)` evaluates to `true`, returning `"0"`.
* **Negative Zero:** Formatting can produce `"-0"`.

### ⚡ Performance & Quality
* **Regex Chaining:** Multiple `.replace()` calls in `normalizeInput` create unnecessary string allocations.
* **Global Pollution:** Functions are attached directly to `window`.
* **Comment Mismatch:** Code comments mention `/g` flag for `toFixed()`, which is misleading.

---

## `index.html`

### 🐞 Bugs & Edge Cases
* **`localStorage` Exceptions:** Lacks a `try...catch` block. Strict privacy modes or Add-in limits can throw exceptions breaking theme initialization.
* **Global Scope Dependency:** Inline `onclick` handlers on symbol buttons might fire before external JS is loaded, throwing `ReferenceError`.
* **Security Vulnerability:** `<a target="_blank">` link is missing `rel="noopener noreferrer"`.

### ⚡ Performance Issues
* **Synchronous Script Loading:** Heavy libraries (Nerdamer, GeoGebra, Chart.js) block parsing. They should use `defer`.
* **Heavy Dependency Lazy Loading:** `deployggb.js` should be lazy-loaded on demand.
* **Inline CSS:** Verbose inline styles degrade caching and bloat HTML payload.

### ♿ Accessibility (a11y)
* **Tab Interface Roles:** Missing standard WAI-ARIA tab attributes (`role="tablist"`, `role="tabpanel"`).
* **Missing Input Labels:** Triangle Calculator labels lack `for` attributes connecting them to inputs.
* **Unlabeled Emoji Button:** Theme toggle button lacks an `aria-label`.
* **Disabled Button Accessibility:** Using the `disabled` attribute hides the `title` explanation from screen readers; `aria-disabled="true"` is better.
* **Decorative SVG:** Loading spinner SVG missing `aria-hidden="true"`.
* **Missing Button Types:** `<button>` elements lack `type="button"`.

### 🧹 Code Quality
* **Separation of Concerns:** Inline JS and `onclick` handlers should be extracted.
* **Inconsistent IDs:** Mix of camelCase and underscore_casing.
* **Redundant ARIA Attributes:** Some inputs use both `<label>` and `aria-label`.
* **Manual Cache Busting:** Query strings (`?v=38`) should be handled by a build tool.

---

## `mathEngine.js`

### 🐞 Bugs & Edge Cases
* **Word Web Content Control Lock Bug:**
  * **Symptom:** Equations inserted via OMML replacement into a `<w:sdt>` (Content Control) become permanently locked in Word Web. The user cannot delete, backspace, or edit the math element.
  * **Root Cause:** Word Web strictly tracks the `paraId` when Enter is pressed. If we try to insert math as plain text, it shares the `paraId` with the line above, causing `insertOoxml("Replace")` to overwrite the previous line. To fix this, we insert a temporary Content Control, which forces Word to generate a unique `paraId`. However, during the OOXML swap, the inner `<w:sdt>` tags are removed from the XML while Word's memory still holds the Content Control object. This mismatch corrupts the Content Control, making it undeletable.
  * **Failed Fixes:** Stripping the `paraId` regex manually causes the overwrite bug again. Calling `cc.delete(true)` before replacing the OOXML causes Word to instantly revert the `paraId` to the duplicate one, also triggering the overwrite bug. Keeping the `<w:sdt>` in the XML but setting `cc.cannotDelete = false` also fails to behave correctly.
  * **Potential "Crappy Solution" (Workaround):** Add a "Slet matematik" (Delete Math) button to the taskpane. Since the user cannot delete the locked box with their keyboard, they can highlight the locked equation and click the button, which calls `context.document.getSelection().clear()` or `delete()` via the Word API to forcefully remove the locked box.
* **Null Reference in Error Handlers:** `errorBox.innerText = error.message;` throws if `errorBox` is null. `resultContainer` and `insertBtn` assumptions in `calculateDraft` also risk exceptions.
* **Comma-to-Dot Conversion Corrupts Vectors:** Replacing `(\d),(\d)` with `.` corrupts input like `gcd(14,21)`.
* **Unsafe Variable Extraction Fallback:** The fallback removing hardcoded function names fails on unsupported functions like `sinh(x)`, incorrectly picking `h` as the variable.
* **Brittle Fallback for Empty Selection:** Truncating text after `:` fails if the math contains ratios (`1:2`) or timestamps.
* **Word API `search()` Special Chars:** Passing raw math expressions to `targetRange.search` fails if they contain special regex/wildcard characters.

### ⚡ Performance Issues
* **Redundant Nerdamer Parsing:** `calculateInline` evaluates math, gets text, and reparses it into Nerdamer to get TeX.
* **Main-thread DOM Parsing Overhead:** Synchronously parsing huge OOXML chunks on the main thread may cause stutter.

### 🧹 Code Quality
* **Severe DRY Violations:** The Word selection narrowing code and complex decimal rounding logic are duplicated heavily.
* **Dead Code:** `if (actionType === 'graph' && base64Image)` is unreachable.
* **Monolithic Function:** `calculateInline` is 300+ lines. The OOXML DOM manipulation should be extracted.

---

## `graphing.js`

### 🐞 Bugs & Edge Cases
* **Aggressive Comma Replacement:** Replaces all commas, breaking GeoGebra 2D coordinates `Point(1,2)`.
* **Fragile String Parsing Fallback:** Truncating text after `:` breaks ratio inputs.
* **Silently Ignored Errors in `evalCommand`:** Ignores the boolean return of `evalCommand`, misleading the user on invalid syntax.
* **Multiple Injections on Timeout:** GeoGebra timeout logic risks duplicate injections if the user clicks twice.
* **Modal Dimension Calculation:** Modal dimension calculation relies on `clientWidth` before the browser finishes rendering the Flexbox layout, causing squished applets.
* **Missing Office JS Context:** Generically catches unhandled `ReferenceError`s if Word API is absent.

### ⚡ Performance Issues
* **Inefficient Polling Loop:** Polling `ggbApplet` via `setInterval` every 100ms instead of awaiting an initialization Promise.
* **Uncached DOM Lookups:** Frequently polled UI functions constantly query `document.getElementById`.

### 🧹 Code Quality
* **Global Namespace Pollution:** Extensively pollutes the window object with GeoGebra state.
* **Inconsistent Variables:** Mixes legacy `var` with `let` and `const`.
* **Magic Numbers:** Numeric literals (15000, 300, 450) scattered instead of centralized constants.
* **Misleading Async Control Flow:** `loadEquationFromWord` relies on internal fail-fast checks instead of proper promise rejection.

---

## `statistics.js`

### 🐞 Bugs & Edge Cases
* **Thousand Separator Logic (`parseData`):** Conversions like `1.234.567` get parsed as `1.234` by `parseFloat`, causing silent data loss.
* **Missing Thousand Separator Handling (`parseGroupedData`):** Fails to strip dots, resulting in `NaN`.
* **Unsorted Groups Breaking Quartiles:** `calculateGrouped` assumes grouped intervals are strictly sorted; fails mathematically if entered out of order.
* **Floating Point Rounding (`statFmt`):** Fails on numbers like `1.005` due to missing Epsilon adjustments.
* **Missing Dependency Checks:** Fails silently if `simple-statistics` or `Chart.js` fails to load.
* **Memory Exhaustion on Large Frequencies:** Expanding intervals to discrete values allocates arrays $O(N)$ size; crashes on frequencies like `1000000`.
* **Decimal Frequencies Corrupt Array Length:** Running a loop `2.5` times causes misalignment of populations.
* **Mixed Interval Data Fallback:** Entering 9 intervals and 1 discrete point forces the script to treat all 10 as discrete points, losing interval precision.
* **`NaN` Output in `statFmt`:** `NaN.toFixed(d)` outputs `"NaN"` straight to the Word document instead of `"N/A"`.

### ⚡ Performance Issues
* **Inefficient Discrete Quantile Calculation:** Expanding midpoints takes $O(N)$ memory and $O(N \log N)$ sorting time.
* **Hardcoded Render Wait:** `await new Promise(r => setTimeout(r, 100))` for Chart.js rendering is a race-condition patch.

### 🧹 Code Quality
* **Regex Redundancy:** `[\s;|\t\n]+` is redundant since `\s` includes tabs and newlines.
* **Hardcoded Office Table Styles:** `"Grid Table 4 - Accent 1"` is localized and breaks in non-English Word versions.
* **Inconsistent Feature Parity:** Ungrouped data gets a boxplot, grouped data generates no charts.

---

## `regression.js`

### 🐞 Bugs & Edge Cases
* **Destructive Word Selection Handling:** `cc.insertText(..., "Replace")` deletes user's highlighted text. Should collapse the range first.
* **Hardcoded Polynomial Degree:** Polynomial regression `FitPoly(L_1, 2)` is locked to quadratic.
* **Data Parsing Flaws:** `replace(/,/g, '.')` mangles thousand separators. Rows with >2 numbers discard extras silently.
* **Incorrect $R^2$ Fallback:** If $R$ is null, it incorrectly displays $R^2 = 0,0000$ instead of "N/A".
* **Race Conditions:** Rapid clicking triggers overlapping async executions, corrupting GeoGebra variables.
* **Fragile XML Parsing:** Assumes highly specific DOM node ancestry.

### ⚡ Performance Issues
* **Expensive DOM Traversal:** `getElementsByTagName("*")` iterates every node in large paragraphs.
* **Whole-Paragraph OOXML Replacement:** Replaces entire paragraph instead of just the Content Control, causing expensive re-parsing in Word.

### 🧹 Code Quality
* **Monolithic Function:** 240-line `calculateRegression` violates Single Responsibility Principle.
* **Missing `DOMParser` Error Validation:** Blindly trusts `xmlDoc` without checking for `<parsererror>`.

---

## `taskpane.css`

### 🐞 Bugs & Edge Cases
* **Hardcoded Focus/Shadow Colors in Dark Mode:** Uses Light Theme's blue shadow hardcoded, failing to adapt in Dark Mode.
* **Insufficient Placeholder Contrast:** `opacity: 0.5` with `--text-muted` fails WCAG contrast ratios in Dark Mode.
* **Missing Keyboard Focus Indicators:** Custom buttons and interactive elements lack `:focus-visible` outlines.
* **Legacy Client Compatibility with `inset`:** `inset: 0` fails on older IE11-based Office Add-in clients.
* **One-way Tab Animations:** Missing "slide out" animations for tab switching.

### ⚡ Responsiveness
* **Rigid Grid Sizing:** `grid-template-columns: 1fr 1fr` forces two columns, squishing inputs on very narrow taskpanes. Use `auto-fit`.
* **Global `overflow-x: hidden`:** Applied to `body`, risking clipped content if users scale text sizes.
* **Fixed Pixel Typography:** Absolute sizes (`13px`) break accessibility scaling. Use relative `rem`.

### 🧹 Code Quality
* **Fragile Anti-Pattern Selectors:** `button:not(.tab-btn):not(.theme-toggle)...` is extremely brittle. Use `.btn` component classes.
* **Overreliance on `!important`:** Forces overrides destructively.
* **Performance Hit from `@import`:** Google Fonts should be in HTML `<link>` to prevent render blocking.
* **Physical vs Logical Properties:** Uses `border-left` instead of `border-inline-start`, complicating future RTL support.

---

## `ommlConverter.js`

### 🐞 Bugs & Critical Logic Errors
* **Race Condition in Initialization:** `initOmmlConverter()` fetches XSLT async, but `latexToOoxml` assumes synchronous availability. Fails if called before network request finishes.
* **Loop Break on Malformed Fractions:** Missing `{` in `\frac` forces a `break`, aborting fraction conversion for the rest of the string.
* **Space Around Slashes Breaks Conversion:** Regex `/[a-zA-Z0-9_.\\{}]/` doesn't match spaces, silently skipping `a / b` conversions.
* **Flawed Parenthesis Wrapping Logic:** `!(num.startsWith('(') && num.endsWith(')'))` blindly avoids wrapping terms like `(x) + (y)`, turning `\frac{(x)+(y)}{z}` into `x + \frac{y}{z}`.

### ⚠️ Edge Cases
* **Regex Fails on Backslashes:** `\\left\(([^()\\]+)\\right\)` skips nested commands like `\left( \frac{1}{2} \right)`.
* **Incomplete Bracket Stripping:** Strips `\left(` but ignores `\left[`, `\left\{`, and `\left|`.
* **Unescaped Braces Break Depth Counters:** Counting `{` and `}` naively breaks when braces are escaped `\{`.
* **XSLT Error Masking:** Masking "Still loading" with "Failed to load".

### ⚡ Performance & Quality
* **Heavy String Manipulation:** Immutable string modification in loops causes O(n²) allocations.
* **Greedy Regex:** `/<math[\s\S]*<\/math>/` causes excessive backtracking.
* **Contradictory Comments:** Stale comments about JS-fallback.
* **AST vs. Regex:** Parsing LaTeX math using regex is fundamentally flawed; should use an AST.

---

## `trigonometry.js`

### 🐞 Bugs & Edge Cases
* **Danish Decimal Comma Ignored:** `parseFloat("3,5")` truncates to `3`. Needs `.replace(',', '.')` before parsing.
* **Floating-Point Precision in SSS:** Math.abs(cosA) <= 1 fails on `1.0000000000000002`. Needs clamping.
* **Degenerate Triangles Bypassing Zero-Check:** Microscopic floats (`1e-14`) bypass `<= 0` checks.

### ⚡ Performance Issues
* **Inefficient Fixed Iteration Loop:** Hardcoded 12 passes for triangle completion; missing early exit condition.
* **No Short-Circuit on Impossible Angles:** Continues iterating after encountering `sinVal > 1`.

### 🧹 Code Quality & UX
* **Missing User Feedback on Insertion Errors:** Failures in `Word.run` log to console but don't inform the UI.
* **Output Precision Comment Mismatch:** Comment says 4 decimals, code enforces `.toFixed(2)`.
* **Redundant Validation Check:** Internal calculations mathematically guarantee `Math.abs(A + B + C - 180) < 0.01`.
