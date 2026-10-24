# Native Office compatibility

This page records what desktop Word, PowerPoint and Excel observed when they opened
OfficeAgent 1.0 output. Every operation `describe_capabilities` advertises has at least one case,
and so do the Word workflows: template population, comparison and assembly. It is evidence for
these files in these applications, not a claim about every Office version, every authoring
application or pixel-identical rendering.

The recorded run used Microsoft 365 Apps build **16.0.20326.20158** (Word, PowerPoint and Excel,
Click-to-Run, 64-bit) on **Windows 11 Pro 10.0.26200**. The exact files Office opened, their
SHA-256 hashes and every observation are in the
[native corpus](../tests/OfficeAgent.Tests/Corpus/v1.0.0/native/README.md).

## Three evidence layers

The layers are kept apart, and none is relabelled as another.

| Layer | How | Runs | What it establishes |
| --- | --- | --- | --- |
| Package and schema | Open XML SDK validator, Office 2019 schema | Every build (`NativeCorpusTests`) | The output adds no schema error its input did not already have. |
| Semantic and part preservation | Open XML assertions and part hashes | Every build | Text, revision marks and authors, and table, slide and chart counts are as asked. Where a case names its allowed parts, no other part changed by a byte. |
| Native | Office object model, [`scripts/native/office_check.ps1`](../scripts/native/office_check.ps1) | On a Windows machine with Office, per release | Office opens the file with repair off and alerts suppressed. What Office itself reports matches the case, including accept and reject in Word. Office saves a copy, and the copy reopens. |
| Visual | PDF exported by Office, rasterised with Poppler | With the native layer, selected cases | Model-assisted inspection of each output beside its own input, for missing objects, overflow and visible corruption. Pagination is not compared with any other engine. |

### Allowed differences

A case passes with these differences, because Office treats them as equivalent:

- **Remapped identifiers.** New paragraph, revision, comment, drawing and relationship ids are
  assigned by the engine. Existing ids are kept.
- **Calculated caches.** Excel recalculates formula results on open. Word recalculates fields
  and page counts. Chart caches are rewritten from the embedded workbook when it is edited.
- **Whitespace markup.** Text the engine writes carries `xml:space="preserve"`. See
  [known cases](#known-unsupported-and-edge-cases) for how this shows up when the input omitted it.
- **Package metadata.** ZIP entry order, compression and timestamps are not preserved, so a
  rerun produces different file bytes with the same parts and meaning.

## Repair detection is proved, not assumed

A harness that cannot see a repair prompt would pass anything. Each run therefore includes one
deliberately corrupt `.docx`, `.pptx` and `.xlsx`. The run fails unless every one of them is
refused, and a control counts only if the same application opened a valid case in the same run.
A worker that fails before Office opens the file does not count as a refusal. Each case runs in
its own process with a time limit, so a modal dialog becomes a recorded failure rather than a
stalled run.

The record published with 1.1.0 did not meet this. Its control workers failed before opening
their files, and those failures were counted as refusals. The current record replaces it: Word,
PowerPoint and Excel each opened their control and refused it, with the refusal recorded in
`results.json`.

The harness was also shown to fail on wrong output. Each tampered copy below was given a
matching manifest hash, so only the checks could catch it:

| Tamper | Result |
| --- | --- |
| Tracked `changeText` output replaced by its input | Fails: 0 revisions, accepted text lacks the replacement |
| Second author of a two-author redline renamed to the first | Fails: revision authors differ |
| `SUM(B2:B3)` rewritten to `SUM(B2:B2)` | Fails: Excel's recalculated value and formula differ |
| Chart's link to its embedded workbook removed | Fails: PowerPoint does not answer (killed at the time limit) |
| Chart's embedded workbook replaced by junk | Fails: PowerPoint refuses to open the deck |

### The record is bound to its files

`results.json` applies only to the corpus beside it, and a test on every build fails otherwise:
the results must name the SHA-256 of the exact `manifest.json` bytes, list exactly the manifest's
case ids once each with the same format and operation, record for every case the output hash
Office opened and that the manifest and the file on disk both carry, pass every individual check,
and account for exactly the corrupt control files, each refused by its own application. Regenerating
the corpus without a new Office run therefore fails the build, and each of those bindings is proved
by a test that tampers a copy of the record. The PDFs reviewed visually are not published; their
hash list is kept with the private release evidence.

## Matrix

<!-- native-matrix:begin -->

83 of 83 cases pass and 3 of 3 corrupt controls are caught, on Office 16.0.20430.20182, Microsoft Windows 11 Pro 10.0.26200. Full hashes are in the corpus manifest.

| Format | Operation | Case | Native checks | Visual | Output SHA-256 |
| --- | --- | --- | --- | --- | --- |
| Word | `changeText` | Tracked replacement of a phrase | 7 passed: revisions at least; revisions after accept all; accepted text contains 'Globex Inc.'; revisions after reject all; rejected text contains 'Acme Corp'; rejected text excludes 'Globex'; Word-saved copy reopens with the same revision count | yes | `35d50dea06ee5cf4…` |
| Word | `revision` | Two authors' tracked edits in successive commits | 9 passed: revision authors; revisions after accept all; accepted text contains 'Globex Inc.'; accepted text contains 'Northwind Traders'; revisions after reject all; rejected text contains 'Acme Corp'; rejected text excludes 'Northwind'; rejected text excludes 'Globex'; Word-saved copy reopens with the same revision count | yes | `2855d831d3ae1907…` |
| Word | `insertParagraphs` | Tracked paragraph insertion | 6 passed: revisions at least; revisions after accept all; accepted text contains 'Inserted obligation paragraph.'; revisions after reject all; rejected text excludes 'Inserted obligation paragraph.'; Word-saved copy reopens with the same revision count |  | `cee070eceaac5df4…` |
| Word | `removeParagraph` | Tracked paragraph removal | 6 passed: revisions at least; revisions after accept all; accepted text excludes 'Service Agreement'; revisions after reject all; rejected text contains 'Service Agreement'; Word-saved copy reopens with the same revision count |  | `01e9d592f6712346…` |
| Word | `format` | Tracked formatting change | 4 passed: revisions at least; revisions after accept all; accepted text contains 'Acme Corp'; Word-saved copy reopens with the same revision count |  | `aca8618bfc955972…` |
| Word | `insert` | Direct paragraph insertion after a found phrase | 3 passed: revisions; text contains 'The Supplier is Acme Corp.'; Word-saved copy reopens with the same revision count |  | `f52b3d216d1e1626…` |
| Word | `revision` | Accept every pending revision | 4 passed: revisions; text contains 'approved'; text excludes 'draft'; Word-saved copy reopens with the same revision count |  | `226c10c440c4de55…` |
| Word | `revision` | Reject every pending revision | 4 passed: revisions; text contains 'draft'; text excludes 'approved'; Word-saved copy reopens with the same revision count |  | `d1120bbc067a3a0d…` |
| Word | `comment` | Comment beside existing review threads | 3 passed: text contains 'Numbered obligation'; comments at least; Word-saved copy reopens with the same revision count | yes | `cef31e6f64b4148b…` |
| Word | `note` | Footnote | 3 passed: footnotes; footnote text; Word-saved copy reopens with the same revision count | yes | `05be8b5b3d1127bc…` |
| Word | `note` | Endnote | 3 passed: endnotes; endnote text; Word-saved copy reopens with the same revision count |  | `3b6facf9b194cd8b…` |
| Word | `fill` | Fill a nested content control | 3 passed: text contains 'Contoso Holdings'; content controls at least; Word-saved copy reopens with the same revision count |  | `47cf36bb2fcc9094…` |
| Word | `insertTable` | Insert a 3x3 table | 5 passed: text contains 'Widget'; tables; table rows; table columns; Word-saved copy reopens with the same revision count | yes | `08bc08a668ac7473…` |
| Word | `insertTableRows` | Append table rows | 4 passed: text contains 'Sprocket'; tables; table rows; Word-saved copy reopens with the same revision count |  | `6be5228e28da19f4…` |
| Word | `removeTableRows` | Remove a table row | 4 passed: text excludes 'Gadget'; tables; table rows; Word-saved copy reopens with the same revision count |  | `f3762ce02d7774b4…` |
| Word | `insertTableColumns` | Insert a table column | 4 passed: text contains 'Total'; tables; table columns; Word-saved copy reopens with the same revision count |  | `216f6baecffc0b43…` |
| Word | `removeTableColumns` | Remove a table column | 4 passed: text excludes 'Price'; tables; table columns; Word-saved copy reopens with the same revision count |  | `5281d665e554cfcc…` |
| Word | `removeTableColumns` | Tracked column removal, resolved in Word | 9 passed: revisions at least; table columns; revisions after accept all; accepted text excludes 'Price'; table columns after accept all; revisions after reject all; rejected text contains 'Price'; table columns after reject all; Word-saved copy reopens with the same revision count |  | `57ffe8b2483f713f…` |
| Word | `repeatTableRow` | Repeat a template row per record | 7 passed: text contains 'Widget'; text contains 'Gadget'; text contains 'Sprocket'; text excludes '{{'; tables; table rows; Word-saved copy reopens with the same revision count | yes | `0e5afbf2fcfd6907…` |
| Word | `mergeTableCells` | Merge the header row's first two cells | 7 passed: text contains 'Item'; text contains 'Qty'; text contains 'Widget'; tables; table rows; cells in the first row; Word-saved copy reopens with the same revision count | yes | `1a1b8dac3627d464…` |
| Word | `removeTable` | Remove a table | 3 passed: text excludes 'Widget'; tables; Word-saved copy reopens with the same revision count |  | `3ccef4b30b3ee94b…` |
| Word | `fillTable` | Refill a table from data, growing it by a row | 8 passed: text contains 'Gizmo'; text contains 'Thingamajig'; text excludes 'Widget'; text excludes 'Gadget'; tables; table rows; table columns; Word-saved copy reopens with the same revision count | yes | `6fa367474f10e631…` |
| Word | `bindData` | Link a phrase to a workbook cell | 4 passed: text contains 'Acme Corp shall provide'; content controls at least; content control tagged 'xl:Summary!B2'; Word-saved copy reopens with the same revision count |  | `48ee8b1e58769dfb…` |
| Word | `bindData` | Link a table to a workbook range | 4 passed: text contains 'Widget'; tables; table alt-text title; Word-saved copy reopens with the same revision count |  | `1f6081e971e61893…` |
| Word | `insertImage` | Insert an inline picture | 2 passed: inline pictures; Word-saved copy reopens with the same revision count | yes | `872286358e617b06…` |
| Word | `removeImage` | Remove the picture | 2 passed: inline pictures; Word-saved copy reopens with the same revision count |  | `e78e91f11bb9d0ac…` |
| Word | `backgroundImage` | Page background picture | 2 passed: header holds a picture behind text; Word-saved copy reopens with the same revision count | yes | `75a6c149ae32ed5f…` |
| Word | `headerFooter` | Header, footer and page number | 3 passed: header contains 'Contoso'; footer contains 'Services agreement'; Word-saved copy reopens with the same revision count | yes | `d0238572ef1f3916…` |
| Word | `pageSetup` | Landscape A4 with wide margins | 2 passed: landscape; Word-saved copy reopens with the same revision count | yes | `d81fd906fabadf31…` |
| Word | `insertBreak` | Page break | 2 passed: pages at least; Word-saved copy reopens with the same revision count | yes | `566cb2a38e67972f…` |
| Word | `setProperty` | Document title property | 2 passed: title property; Word-saved copy reopens with the same revision count |  | `297926bb70e52378…` |
| Word | `setProperty` | Compatibility mode set to Word 2010 | 2 passed: compatibility mode; Word-saved copy reopens with the same revision count |  | `d8af79e566108c7f…` |
| Word | `link` | Hyperlink on existing text | 6 passed: revisions; hyperlinks; hyperlink address; hyperlink text; Word reads the same text as in the input; Word-saved copy reopens with the same revision count |  | `50e0adfc2af69fa3…` |
| Word | `format` | A paragraph attached to the document's existing list | 2 passed: list numbers in document order; Word-saved copy reopens with the same revision count |  | `4795d36119360f1f…` |
| Word | `defineStyle` | Define and apply a heading style | 2 passed: style 'Contract Heading' exists; Word-saved copy reopens with the same revision count |  | `d0f89a75ee49280b…` |
| Word | `clearStyles` | Clear direct formatting on a bolded phrase | 3 passed: revisions; Word reads the same text as in the input; Word-saved copy reopens with the same revision count |  | `3ccef4b30b3ee94b…` |
| Word | `clearStyles` | Clear formatting on runs written without xml:space | 2 passed: text contains 'Quarterly target'; Word-saved copy reopens with the same revision count |  | `14083271bed742cd…` |
| Word | `copyStyles` | Copy direct formatting between paragraphs | 2 passed: Word reads the same text as in the input; Word-saved copy reopens with the same revision count |  | `c2713a3eeca90607…` |
| Word | `create` | Blank document from OfficeAgent, then written | 4 passed: revisions; text contains 'Created by OfficeAgent.'; Word compatibility mode; Word-saved copy reopens with the same revision count |  | `7f5a618d853a4d64…` |
| Word | `template` | Template with a slot and a repeating row, populated | 7 passed: text contains 'Fabrikam Ltd'; text contains 'Consulting'; text contains 'Support'; text excludes '{{'; text excludes 'CUSTOMER'; table rows; Word-saved copy reopens with the same revision count | yes | `9b2665374e99cc25…` |
| Word | `template` | Template image slot bound through a template batch on the filesystem provider | 5 passed: text contains 'Quote for Fabrikam Ltd'; text excludes 'CUSTOMER'; picture size in points; inline pictures; Word-saved copy reopens with the same revision count | yes | `5f9fdb8f64d3df91…` |
| Word | `comparison` | Comparison redline of a table-cell edit | 10 passed: revisions at least; tables; table rows; revisions after accept all; accepted text contains 'Gizmo'; accepted text excludes 'Gadget'; revisions after reject all; rejected text contains 'Gadget'; rejected text excludes 'Gizmo'; Word-saved copy reopens with the same revision count | yes | `26a79c6435c109b6…` |
| Word | `comparison` | Comparison redline between two versions | 7 passed: revisions at least; revisions after accept all; accepted text contains 'Globex Inc.'; revisions after reject all; rejected text contains 'Acme Corp'; rejected text excludes 'Globex'; Word-saved copy reopens with the same revision count | yes | `18e5c418ecceb31a…` |
| Word | `assembly` | Two documents assembled, each keeping its header | 6 passed: revisions; text contains 'Acme Corp'; text contains 'Appendix A: rate card.'; sections at least; Word compatibility mode; Word-saved copy reopens with the same revision count | yes | `e0e70d7c00e4ef0e…` |
| PowerPoint | `insertSlide` | Two slides added to a blank deck | 4 passed: slides; text contains 'Second'; text contains 'Closing point'; PowerPoint-saved copy reopens with the same slide count | yes | `b33d5a8788982512…` |
| PowerPoint | `removeSlide` | Remove a slide | 3 passed: slides; text excludes 'Another point'; PowerPoint-saved copy reopens with the same slide count |  | `a21e4f3f5e04f7d9…` |
| PowerPoint | `moveSlide` | Move the last slide first | 3 passed: slides; first slide contains 'Third'; PowerPoint-saved copy reopens with the same slide count |  | `d99005ad837297b7…` |
| PowerPoint | `duplicateSlide` | Duplicate a slide | 2 passed: slides; PowerPoint-saved copy reopens with the same slide count |  | `2939f1f558887865…` |
| PowerPoint | `section` | Add a section | 3 passed: sections at least; section 'Financials'; PowerPoint-saved copy reopens with the same slide count |  | `dbc7558a49a55d58…` |
| PowerPoint | `transition` | Fade transition on every slide | 3 passed: slides; every slide has a transition; PowerPoint-saved copy reopens with the same slide count |  | `b7c3fae97f62905e…` |
| PowerPoint | `animate` | Fade-in animation on a shape | 2 passed: animations at least; PowerPoint-saved copy reopens with the same slide count |  | `07379b11240d53a6…` |
| PowerPoint | `headerFooter` | Footer and slide numbers | 2 passed: footer contains 'Confidential'; PowerPoint-saved copy reopens with the same slide count | yes | `1534e5ec30ad5a84…` |
| PowerPoint | `insertShape` | Insert a text box | 2 passed: text contains 'Accent'; PowerPoint-saved copy reopens with the same slide count | yes | `be284d2d3352bae1…` |
| PowerPoint | `removeShape` | Remove the text box | 2 passed: text excludes 'Accent'; PowerPoint-saved copy reopens with the same slide count |  | `92e7e1d410e8c631…` |
| PowerPoint | `insertChart` | Native column chart with embedded data | 4 passed: charts; chart data opens for editing; chart series values; PowerPoint-saved copy reopens with the same slide count | yes | `3b1e2e493640a2bd…` |
| PowerPoint | `updateChart` | Replace the chart's data | 4 passed: charts; chart data opens for editing; chart series values; PowerPoint-saved copy reopens with the same slide count | yes | `747dc7a27ecaab3e…` |
| PowerPoint | `setChartData` | New data in an existing chart, growing it to two series | 5 passed: charts; chart data opens for editing; chart series values; chart series; PowerPoint-saved copy reopens with the same slide count | yes | `f67e9687d5b2075f…` |
| PowerPoint | `template` | Template chart slot bound through a template batch on the filesystem provider | 4 passed: charts; chart data opens for editing; chart series values; PowerPoint-saved copy reopens with the same slide count | yes | `370307b757c33840…` |
| PowerPoint | `insertImage` | Insert a picture | 2 passed: pictures; PowerPoint-saved copy reopens with the same slide count | yes | `9d491c9187de3806…` |
| PowerPoint | `removeImage` | Remove the picture | 2 passed: pictures; PowerPoint-saved copy reopens with the same slide count |  | `855081e05e191c78…` |
| PowerPoint | `backgroundImage` | Slide background picture | 2 passed: slide background is a picture; PowerPoint-saved copy reopens with the same slide count | yes | `b9502e3ba67bbbb4…` |
| PowerPoint | `insertMedia` | Embedded video with a poster frame | 2 passed: media objects; PowerPoint-saved copy reopens with the same slide count |  | `55c08b3bab3dfadb…` |
| PowerPoint | `insertTable` | Insert a table | 3 passed: text contains 'EMEA'; tables; PowerPoint-saved copy reopens with the same slide count | yes | `da8de190107cc34c…` |
| PowerPoint | `insertTableRows` | Append a table row | 3 passed: text contains 'APAC'; tables; PowerPoint-saved copy reopens with the same slide count |  | `5f201f0a46a18ca5…` |
| PowerPoint | `removeTableRows` | Remove a table row | 2 passed: tables; PowerPoint-saved copy reopens with the same slide count |  | `3fe8e7f68a0f676f…` |
| PowerPoint | `insertTableColumns` | Insert a table column | 3 passed: text contains 'Q2'; tables; PowerPoint-saved copy reopens with the same slide count |  | `2b2a212234ecc863…` |
| PowerPoint | `removeTableColumns` | Remove a table column | 2 passed: tables; PowerPoint-saved copy reopens with the same slide count |  | `0ca22ddca29dd15c…` |
| PowerPoint | `removeTable` | Remove the table | 2 passed: tables; PowerPoint-saved copy reopens with the same slide count |  | `4c05b7f2b7df87d1…` |
| PowerPoint | `fillTable` | Refill a table from data, growing it by a row | 5 passed: text contains 'North'; text contains 'South'; text excludes 'EMEA'; tables; PowerPoint-saved copy reopens with the same slide count | yes | `2295b6f079462303…` |
| PowerPoint | `bindData` | Link a shape to a workbook cell | 3 passed: slides; shape named 'xl:Revenue'; PowerPoint-saved copy reopens with the same slide count |  | `c1d78d7f517498e2…` |
| PowerPoint | `changeText` | Replace the title text | 2 passed: text contains 'Revised title'; PowerPoint-saved copy reopens with the same slide count |  | `9706c68c41b457fc…` |
| PowerPoint | `insert` | Insert a paragraph after the title | 2 passed: text contains 'Draft for discussion'; PowerPoint-saved copy reopens with the same slide count |  | `73fc9ad8db2dac4c…` |
| PowerPoint | `format` | Bold, coloured title | 2 passed: text contains 'Second'; PowerPoint-saved copy reopens with the same slide count | yes | `151b568b066c2bf0…` |
| PowerPoint | `clearStyles` | Clear direct formatting on the title | 2 passed: text contains 'Second'; PowerPoint-saved copy reopens with the same slide count |  | `e0e204265272eb5c…` |
| PowerPoint | `copyStyles` | Copy paragraph formatting between lines | 2 passed: slides; PowerPoint-saved copy reopens with the same slide count |  | `e0e204265272eb5c…` |
| PowerPoint | `comment` | Modern comment on a slide | 2 passed: comments at least; PowerPoint-saved copy reopens with the same slide count |  | `ad83461fc9588193…` |
| PowerPoint | `setProperty` | Deck title property | 2 passed: title property; PowerPoint-saved copy reopens with the same slide count |  | `16312eb742b0df0b…` |
| PowerPoint | `fill` | Fill a named template shape | 3 passed: text contains 'Northwind Traders Limited'; text excludes '[CLIENT]'; PowerPoint-saved copy reopens with the same slide count | yes | `e94fc73aef8d57c7…` |
| Excel | `setCell` | Text, numbers and a SUM formula | 6 passed: no Excel repair log; A1 value; B2 value; B4 value; B4 formula; Excel-saved copy reopens | yes | `f05e79d5b562d251…` |
| Excel | `calculation` | 107 allowlisted formulas calculated by the engine, recalculated by Excel | 3 passed: no Excel repair log; formulas where Excel agrees with the cached value; Excel-saved copy reopens |  | `bc781da55f6bcc55…` |
| Excel | `setProperty` | Workbook title property | 3 passed: title property; no Excel repair log; Excel-saved copy reopens |  | `e3038e75c399abf0…` |
| Excel | `comment` | Cell comment | 3 passed: no Excel repair log; comments at least; Excel-saved copy reopens |  | `7a64ee364ad05ac3…` |
| Excel | `appendTableRows` | Append rows to an Excel table | 5 passed: no Excel repair log; cells contains 'APAC'; tables; table data rows grew by; Excel-saved copy reopens | yes | `c21ecabdc2453d3a…` |

<!-- native-matrix:end -->

## Known unsupported and edge cases

### Refused before anything is written

Each of these is refused with a stable code. The input keeps its exact bytes, both in memory and
through the filesystem provider, and no file appears beside it
(`Unsupported_cases_leave_the_stored_file_and_folder_untouched`).

| Case | Code |
| --- | --- |
| `setCell` on one member of an Excel shared-formula group | `invalid-operation` (expand the group first) |
| A tracked edit in PowerPoint, which has no tracked-change vocabulary | `invalid-operation` |
| `fill` naming a shape the deck does not have | `anchor-not-found` |
| `changeText` whose expected text is no longer in the paragraph | `expect-mismatch` |
| `removeParagraph` on a paragraph id that does not exist | `anchor-not-found` |

### Behaviour to expect in Office

- **Compatibility mode follows the input.** A Word package with no declared compatibility mode
  opens as a Word 2007 document ("Compatibility Mode"). OfficeAgent keeps whatever mode the input
  has and never upgrades a document silently. Documents it creates declare mode 15 (Word 2013 and
  later), and so does an assembly whose first source does. Before 1.0 a created document had no
  mode, and Word opened it in compatibility mode.
- **Whitespace written without `xml:space="preserve"`.** Word hides leading and trailing spaces
  in such text. Text OfficeAgent rewrites carries the attribute, so a space another producer wrote
  that way, and Word hid, becomes visible once that run is edited. Word writes the attribute
  wherever a space matters, so this arises with other producers' files.
- **Tables without set widths autofit.** Word sizes such a table's columns to its content, so
  filling a template row with shorter or longer values changes the column widths. The table
  markup, including its grid, is unchanged.
- **A tracked column removal stays visible until accepted.** Word shows the column, marked
  deleted, until the revision is accepted. The corpus accepts it (2 columns) and rejects it
  (3 columns) in Word.
- **A template image is drawn at the size the binding gives.** `TemplateImageValue.WidthPx`
  and `HeightPx` default to 200 by 200, whatever the picture's own shape, so a binding that
  omits them shows a 2:1 photograph square. Set both to the picture's proportions. The corpus
  binds 240 by 120 and Word reports 180 by 90 points.
- **Displayed numbers follow the machine's locale.** Excel showed `12,5` for 12.5 on the recorded
  machine. Values were compared through `Value2`, which is locale-independent.

### What this evidence does not cover

- Other Office builds, Office for Mac, Office on the web and mobile. Only the build above was run.
- Other authoring applications. LibreOffice rendering is covered separately by the
  [renderer reference](../deploy/renderer/README.md).
- **Silent repair in Word and PowerPoint.** Neither application exposes a repair log through
  automation. The harness detects a refused open, and each case's checks detect lost content, but
  a repair that kept everything checked would pass. Excel's repair log is checked on every open,
  but the corrupt workbook control was refused rather than repaired, so no run has yet shown that
  check catching a repair.
- **Redline appearance.** Word exports tracked documents without markup under automation, so
  the PDFs show text, not revision marks. Revisions are verified through Word's object model,
  including accept and reject, not visually.
- Pixel-identical layout, or pagination compared with any other engine.

## Reproducing the run

```powershell
$env:OFFICEAGENT_NATIVE_CORPUS_OUTPUT = 'C:\temp\officeagent-native'
dotnet test tests/OfficeAgent.Tests --filter "FullyQualifiedName~NativeCorpusTests"
Remove-Item Env:OFFICEAGENT_NATIVE_CORPUS_OUTPUT
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/native/office_check.ps1 -Corpus C:\temp\officeagent-native
```

The harness drives Word, PowerPoint and Excel through COM and closes them when it finishes. It
kills any instance of an application that stops answering, so run it where no one is using Office.
It writes `results.json` beside the manifest, Office-saved copies to `roundtrip/`, and PDFs of
each visual case and of its input to `pdf/`.
