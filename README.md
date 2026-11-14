# TR-181 Data Model: XML to Excel Converter

A Python script that turns the Broadband Forum's TR-181 data model (published as a large, deeply nested XML file) into a flat Excel sheet with one row per object or parameter. Built while testing telecom gateways at Intellica, where QA engineers needed to compare device output against the spec without reading raw XML.

## What it does

- Parses the official `tr-181-*-cwmp-full.xml` file
- Walks every object and parameter in the model
- Resolves data types that are defined by reference (type A points to type B, which points to type C) so each row shows the final concrete type, size limits and enum values
- Expands the spec's description macros (`{{param}}`, `{{object}}`, `{{reference}}`, `{{bibref}}`, `{{empty}}`, `{{numentries}}`, `{{pattern}}`) into readable text
- Optionally joins the XML with the HTML version of the spec to pull in the human-readable descriptions
- Writes an `.xlsx` with object rows highlighted so the hierarchy is easy to scan

## Output columns

| Column | Meaning |
|---|---|
| Object Name | Parent object path, e.g. `Device.DeviceInfo.` |
| Parameter Name | Parameter name, empty for object rows |
| Full Path | `Object Name` + `Parameter Name` |
| Description | Resolved description text |
| Data Type | Concrete type after reference resolution, e.g. `string(64)`, `unsignedInt`, `boolean` |
| Object Default | Default value where the spec defines one |
| Is Object | `True` for object rows, `False` for parameters |
| Access | `readOnly` or `readWrite` |
| Version | Spec version the entry was introduced in |

Example rows (shortened):

| Object Name | Parameter Name | Full Path | Data Type | Access |
|---|---|---|---|---|
| `Device.DeviceInfo.` | | `Device.DeviceInfo.` | | |
| `Device.DeviceInfo.` | `Manufacturer` | `Device.DeviceInfo.Manufacturer` | `string(64)` | readOnly |
| `Device.DeviceInfo.` | `SoftwareVersion` | `Device.DeviceInfo.SoftwareVersion` | `string(64)` | readOnly |
| `Device.WiFi.SSID.{i}.` | `SSID` | `Device.WiFi.SSID.{i}.SSID` | `string(32)` | readWrite |

## Requirements

- Python 3.9 or newer
- `pandas`, `openpyxl`, `beautifulsoup4`

```bash
pip install -r requirements.txt
```

## Usage

1. Download the TR-181 data model from the Broadband Forum's CWMP data model page. The script expects the "full" CWMP XML, for example `tr-181-2-19-0-cwmp-full.xml`.
2. Put the XML file in the same folder as the script.
3. Optional: put the HTML version of the same spec (`tr-181-*.html`) in the same folder. If present, descriptions are taken from it; if not, descriptions come from the XML.
4. Run:

```bash
python xml-to-excel-converter-3.py
```

5. Open `TR-181-DM (Final).xlsx` in the same folder.

The input file name and output file name are set near the bottom of `main()`. Change them there if you are using a different spec version.

## How the hard parts work

**Type references.** In the XML, a parameter rarely says its type directly. It says `<dataType ref="Alias"/>`, and `Alias` is defined elsewhere as `<dataType base="string"><size maxLength="64"/></dataType>`. Some references chain several levels deep. `resolve_datatype_reference()` follows the chain recursively, keeping a `visited` set so a bad spec cannot loop forever, and merges size limits and enumerations from each level into one final type string.

**Templates.** Shared parameter definitions live in `<template>` elements that can inherit from other templates. The script loads all templates first, then applies inheritance (child values win, parent values fill gaps), then uses the result when a parameter references a template.

**Description macros.** Spec descriptions contain placeholders like `{{param|Enable}}` or `{{object}}`. `substitute_macros()` rewrites these into plain text using the current parameter's context so the Excel cell reads like a sentence.

**HTML join.** The HTML spec has cleaner descriptions than the XML. `main()` walks the HTML table rows, tracks the current object path, and builds a `full path -> description` map. `normalize_path()` makes the XML and HTML paths comparable (same handling of `{i}` instance markers and trailing dots).

## Limitations

- Single script, no CLI flags. File names are hard-coded in `main()`.
- No unit tests yet.
- Built against TR-181 Issue 2 Amendment 19. Other versions should work but have not been verified.
