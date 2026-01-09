# Style Factory — Unified ODS Formatting for Trade Control

## Purpose

The Style Factory provides a unified formatting engine for all ODS document generators in Trade Control.  
It converts semantic formatting codes (e.g., `NUM2`, `PCT1`, `CASH2_NEG_CELL`) into fully formed ODF number styles and table‑cell styles, then injects them into the final `.ods` file.

This subsystem exists because:
- ODF Python libraries do not fully implement the ODF number‑style specification.
- LibreOffice Calc requires precise XML structures for correct formatting.
- Formatting logic must be consistent and reusable across generators.

The Style Factory isolates this complexity so generators express intent, not implementation.

## High‑Level Architecture

Generator → Semantic Style → Mapping Engine → XML Injector → Final ODS

### 1. Semantic Layer (`semantic/`)

Defines the vocabulary of formatting used across the ERP.

- Maps SQL template codes to canonical cell style names.
- Encodes meaning, not rendering.
- Canonical examples:
  - Numbers: `NUM0_CELL`, `NUM1_CELL`, `NUM2_CELL`
  - Percentages: `PCT0_CELL`, `PCT1_CELL`, `PCT2_CELL`
  - Accounting cash: `CASH2_POS_CELL`, `CASH2_NEG_CELL` (optional base `CASH2_CELL`)
  - Text: `TEXT_CELL`

Generators assign these names to cells via `table:style-name`.

Invariant: Semantic names must remain stable across the system.

### 2. Mapping Layer (`mapping/`)

Interprets semantic names into formatting instructions.

Responsibilities:
- Determine decimal places and format type (number, percentage, cash).
- Apply accounting rules:
  - parentheses for negatives
  - attempt to suppress leading minus via `number:display-factor="-1"`
  - red text for negative cash via cell style
- Produce internal specs for data‑styles and cell‑styles.

Invariant: Mapping must be deterministic and independent of generator logic.

### 3. Rendering Layer (`rendering/`)

Transforms mapping results into actual ODF XML.

Responsibilities:
- Extract `content.xml` from the `.ods` archive.
- Scan for used semantic styles.
- Generate and inject into `office:automatic-styles`:
  - `<number:number-style>` (grouping, decimal places)
  - `<number:percentage-style>` (decimal places + trailing `%`)
  - `<style:style family="table-cell">` (binds `style:data-style-name`, sets red text via `fo:color`)
  - Optionally `style:map` for base CASH styles (route POS/NEG by condition)
- Ensure each numeric `table:table-cell` includes a `text:p` child to avoid display quirks.
- Optionally strip default sheet artifacts (e.g., `Feuille1`).
- Repackage the ODS without losing entries.

Invariant: Rendered XML must conform to LibreOffice’s interpretation of ODF 1.2 and be idempotent.

### 4. Public API (`engine.py`)

Single entry points:

```python
from style_factory import apply_styles_bytes

updated = apply_styles_bytes(ods_bytes, locale=("en","GB"), strip_defaults=True)
```
Generators call this after writing structural content with **odfdo**.

**Invariant**: Generators must not attempt to create number formats themselves.

## Design Principles

- Separation of concerns:
  - generators: data + layout
  - style factory: formatting
  - odfdo: structure
  - lxml: precise XML injection
- Semantic first: formatting expressed as meaning.
- Minimal coupling: generators avoid ODF internals.
- Deterministic output: same semantic input → same XML output.
- Extensibility: add new formats by extending semantic definitions, mapping rules, or rendering templates without modifying generators.
- Idempotency: multiple runs produce stable XML without duplicates.

## Usage Contract for Generators

1. Assign semantic style names to cells:

```python
# odfdo
cell.set_attribute("table:style-name", "NUM2_CELL") 
cell.set_attribute("table:style-name", "CASH2_NEG_CELL")

# or use a base accounting style with conditional maps
cell.set_attribute("table:style-name", "CASH2_CELL")
```
2. Save the `.ods` using odfdo.

3. Call the Style Factory:

```python
from style_factory import apply_styles_bytes 

updated = apply_styles_bytes(ods_bytes, locale=("en","GB"), strip_defaults=True)
```
4. Do **not** attempt to create number formats manually.

## Accounting formats

- •	Negative styles render with parentheses and red text. We set number:display-factor="-1" in the negative data‑style to suppress the leading minus (LibreOffice may still show a minus in some builds).
- Note: Some LibreOffice builds may still show a leading minus. Using a base `CASHx_CELL` with `style:map` to route values to `CASHx_POS_CELL` or `CASHx_NEG_CELL` ensures consistent POS/NEG application without formulas.

## Locale defaults

- The Style Factory updates:
    - styles.xml default styles (paragraph, text, table-cell) to a single style:text-properties block using the requested locale.
    - content.xml data styles to include number:language and number:country on the inner number elements.
    - meta.xml dc:language to the requested locale.
- Legacy number:currency-style entries from base templates are stripped to avoid unintended locale fallback.
- Configurable via `locale=("en","GB")` in the public API.

## Why This Exists

ODF number formatting is powerful but under‑implemented in Python libraries.
LibreOffice Calc requires:

- explicit number‑styles
- explicit cell‑styles
- correct namespace usage
- correct ordering within `office:automatic-styles`
- correct handling of negative accounting formats
- correct and consistent locale stamping across styles.xml, content.xml, and meta.xml
 
The Style Factory ensures these requirements are met consistently and automatically.

## Future Extensions

- Locale‑aware currency symbols
- Thousands‑separator variations
- Date/time formats
- Style caching
- Multi‑sheet style analysis
