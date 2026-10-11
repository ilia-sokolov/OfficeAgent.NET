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

77 of 77 cases pass and 3 of 3 corrupt controls are caught, on Office 16.0.20430.20118, Microsoft Windows 11 Pro 10.0.26200. Full hashes are in the corpus manifest.

| Format | Operation | Case | Native checks | Visual | Output SHA-256 |
| --- | --- | --- | --- | --- | --- |
| Word | `changeText` | Tracked replacement of a phrase | 7 passed: revisions at least; revisions after accept all; accepted text contains 'Globex Inc.'; revisions after reject all; rejected text contains 'Acme Corp'; rejected text excludes 'Globex'; Word-saved copy reopens with the same revision count | yes | `b4f82476ed84f11f…` |
| Word | `revision` | Two authors' tracked edits in successive commits | 9 passed: revision authors; revisions after accept all; accepted text contains 'Globex Inc.'; accepted text contains 'Northwind Traders'; revisions after reject all; rejected text contains 'Acme Corp'; rejected text excludes 'Northwind'; rejected text excludes 'Globex'; Word-saved copy reopens with the same revision count | yes | `0a06e85a379598fb…` |
| Word | `insertParagraphs` | Tracked paragraph insertion | 6 passed: revisions at least; revisions after accept all; accepted text contains 'Inserted obligation paragraph.'; revisions after reject all; rejected text excludes 'Inserted obligation paragraph.'; Word-saved copy reopens with the same revision count |  | `d786ecc98fe00bbd…` |
| Word | `removeParagraph` | Tracked paragraph removal | 6 passed: revisions at least; revisions after accept all; accepted text excludes 'Service Agreement'; revisions after reject all; rejected text contains 'Service Agreement'; Word-saved copy reopens with the same revision count |  | `d5096efec398926d…` |
| Word | `format` | Tracked formatting change | 4 passed: revisions at least; revisions after accept all; accepted text contains 'Acme Corp'; Word-saved copy reopens with the same revision count |  | `2134013b58ff83e5…` |
| Word | `insert` | Direct paragraph insertion after a found phrase | 3 passed: revisions; text contains 'The Supplier is Acme Corp.'; Word-saved copy reopens with the same revision count |  | `5de6535d508476a0…` |
| Word | `revision` | Accept every pending revision | 4 passed: revisions; text contains 'approved'; text excludes 'draft'; Word-saved copy reopens with the same revision count |  | `226c10c440c4de55…` |
| Word | `revision` | Reject every pending revision | 4 passed: revisions; text contains 'draft'; text excludes 'approved'; Word-saved copy reopens with the same revision count |  | `d1120bbc067a3a0d…` |
| Word | `comment` | Comment beside existing review threads | 3 passed: text contains 'Numbered obligation'; comments at least; Word-saved copy reopens with the same revision count | yes | `cddcaea3d58876a4…` |
| Word | `note` | Footnote | 3 passed: footnotes; footnote text; Word-saved copy reopens with the same revision count | yes | `84f99344cc73d84d…` |
| Word | `note` | Endnote | 3 passed: endnotes; endnote text; Word-saved copy reopens with the same revision count |  | `5509a3373fc624e9…` |
| Word | `fill` | Fill a nested content control | 3 passed: text contains 'Contoso Holdings'; content controls at least; Word-saved copy reopens with the same revision count |  | `47cf36bb2fcc9094…` |
| Word | `insertTable` | Insert a 3x3 table | 5 passed: text contains 'Widget'; tables; table rows; table columns; Word-saved copy reopens with the same revision count | yes | `08ada7f8a709ae2a…` |
| Word | `insertTableRows` | Append table rows | 4 passed: text contains 'Sprocket'; tables; table rows; Word-saved copy reopens with the same revision count |  | `2cecb5350d7434ed…` |
| Word | `removeTableRows` | Remove a table row | 4 passed: text excludes 'Gadget'; tables; table rows; Word-saved copy reopens with the same revision count |  | `68d58eb714555fb3…` |
| Word | `insertTableColumns` | Insert a table column | 4 passed: text contains 'Total'; tables; table columns; Word-saved copy reopens with the same revision count |  | `9ec6307dbf097124…` |
| Word | `removeTableColumns` | Remove a table column | 4 passed: text excludes 'Price'; tables; table columns; Word-saved copy reopens with the same revision count |  | `de73cfe35bb516b1…` |
| Word | `removeTableColumns` | Tracked column removal, resolved in Word | 9 passed: revisions at least; table columns; revisions after accept all; accepted text excludes 'Price'; table columns after accept all; revisions after reject all; rejected text contains 'Price'; table columns after reject all; Word-saved copy reopens with the same revision count |  | `c24c73f55aa950e4…` |
| Word | `repeatTableRow` | Repeat a template row per record | 7 passed: text contains 'Widget'; text contains 'Gadget'; text contains 'Sprocket'; text excludes '{{'; tables; table rows; Word-saved copy reopens with the same revision count | yes | `16f96d58e3e27276…` |
| Word | `mergeTableCells` | Merge the header row's first two cells | 7 passed: text contains 'Item'; text contains 'Qty'; text contains 'Widget'; tables; table rows; cells in the first row; Word-saved copy reopens with the same revision count | yes | `4ab6b7d00aed77a1…` |
| Word | `removeTable` | Remove a table | 3 passed: text excludes 'Widget'; tables; Word-saved copy reopens with the same revision count |  | `38ae3137531e9287…` |
| Word | `insertImage` | Insert an inline picture | 2 passed: inline pictures; Word-saved copy reopens with the same revision count | yes | `efdec8fc55dd4f3d…` |
| Word | `removeImage` | Remove the picture | 2 passed: inline pictures; Word-saved copy reopens with the same revision count |  | `628145f3bb28f096…` |
| Word | `backgroundImage` | Page background picture | 2 passed: header holds a picture behind text; Word-saved copy reopens with the same revision count | yes | `540103ff024f9928…` |
| Word | `headerFooter` | Header, footer and page number | 3 passed: header contains 'Contoso'; footer contains 'Services agreement'; Word-saved copy reopens with the same revision count | yes | `1cc2369af1188c1d…` |
| Word | `pageSetup` | Landscape A4 with wide margins | 2 passed: landscape; Word-saved copy reopens with the same revision count | yes | `0cceb427993f7371…` |
| Word | `insertBreak` | Page break | 2 passed: pages at least; Word-saved copy reopens with the same revision count | yes | `b2cf274bdb2e7434…` |
| Word | `setProperty` | Document title property | 2 passed: title property; Word-saved copy reopens with the same revision count |  | `113fb26636f1781b…` |
| Word | `setProperty` | Compatibility mode set to Word 2010 | 2 passed: compatibility mode; Word-saved copy reopens with the same revision count |  | `f8f7d92289160366…` |
| Word | `link` | Hyperlink on existing text | 6 passed: revisions; hyperlinks; hyperlink address; hyperlink text; Word reads the same text as in the input; Word-saved copy reopens with the same revision count |  | `e2ee0d4e1fe9c708…` |
| Word | `format` | A paragraph attached to the document's existing list | 2 passed: list numbers in document order; Word-saved copy reopens with the same revision count |  | `548e2efe8fd0aaae…` |
| Word | `defineStyle` | Define and apply a heading style | 2 passed: style 'Contract Heading' exists; Word-saved copy reopens with the same revision count |  | `67194130e1cb209c…` |
| Word | `clearStyles` | Clear direct formatting on a bolded phrase | 3 passed: revisions; Word reads the same text as in the input; Word-saved copy reopens with the same revision count |  | `38ae3137531e9287…` |
| Word | `clearStyles` | Clear formatting on runs written without xml:space | 2 passed: text contains 'Quarterly target'; Word-saved copy reopens with the same revision count |  | `14083271bed742cd…` |
| Word | `copyStyles` | Copy direct formatting between paragraphs | 2 passed: Word reads the same text as in the input; Word-saved copy reopens with the same revision count |  | `c2713a3eeca90607…` |
| Word | `create` | Blank document from OfficeAgent, then written | 4 passed: revisions; text contains 'Created by OfficeAgent.'; Word compatibility mode; Word-saved copy reopens with the same revision count |  | `b866358ea6e2c978…` |
| Word | `template` | Template with a slot and a repeating row, populated | 7 passed: text contains 'Fabrikam Ltd'; text contains 'Consulting'; text contains 'Support'; text excludes '{{'; text excludes 'CUSTOMER'; table rows; Word-saved copy reopens with the same revision count | yes | `f2f72bb1017cd5f7…` |
| Word | `template` | Template image slot bound through a template batch on the filesystem provider | 5 passed: text contains 'Quote for Fabrikam Ltd'; text excludes 'CUSTOMER'; picture size in points; inline pictures; Word-saved copy reopens with the same revision count | yes | `3f3bee9bfa7f3e36…` |
| Word | `comparison` | Comparison redline of a table-cell edit | 10 passed: revisions at least; tables; table rows; revisions after accept all; accepted text contains 'Gizmo'; accepted text excludes 'Gadget'; revisions after reject all; rejected text contains 'Gadget'; rejected text excludes 'Gizmo'; Word-saved copy reopens with the same revision count | yes | `4ab48981a8249103…` |
| Word | `comparison` | Comparison redline between two versions | 7 passed: revisions at least; revisions after accept all; accepted text contains 'Globex Inc.'; revisions after reject all; rejected text contains 'Acme Corp'; rejected text excludes 'Globex'; Word-saved copy reopens with the same revision count | yes | `f9db52becd702fb9…` |
| Word | `assembly` | Two documents assembled, each keeping its header | 6 passed: revisions; text contains 'Acme Corp'; text contains 'Appendix A: rate card.'; sections at least; Word compatibility mode; Word-saved copy reopens with the same revision count | yes | `b1a3d9df36d0b21d…` |
| PowerPoint | `insertSlide` | Two slides added to a blank deck | 4 passed: slides; text contains 'Second'; text contains 'Closing point'; PowerPoint-saved copy reopens with the same slide count | yes | `e38b0612a48b1c01…` |
| PowerPoint | `removeSlide` | Remove a slide | 3 passed: slides; text excludes 'Another point'; PowerPoint-saved copy reopens with the same slide count |  | `685b7fce50798742…` |
| PowerPoint | `moveSlide` | Move the last slide first | 3 passed: slides; first slide contains 'Third'; PowerPoint-saved copy reopens with the same slide count |  | `a6e54534cf9128a0…` |
| PowerPoint | `duplicateSlide` | Duplicate a slide | 2 passed: slides; PowerPoint-saved copy reopens with the same slide count |  | `7df863ff27ad68c4…` |
| PowerPoint | `section` | Add a section | 3 passed: sections at least; section 'Financials'; PowerPoint-saved copy reopens with the same slide count |  | `a4b13d15ae661c7e…` |
| PowerPoint | `transition` | Fade transition on every slide | 3 passed: slides; every slide has a transition; PowerPoint-saved copy reopens with the same slide count |  | `dd01dbbdd4098907…` |
| PowerPoint | `animate` | Fade-in animation on a shape | 2 passed: animations at least; PowerPoint-saved copy reopens with the same slide count |  | `6b7cf544fd19a09f…` |
| PowerPoint | `headerFooter` | Footer and slide numbers | 2 passed: footer contains 'Confidential'; PowerPoint-saved copy reopens with the same slide count | yes | `5a42af8e174f90e8…` |
| PowerPoint | `insertShape` | Insert a text box | 2 passed: text contains 'Accent'; PowerPoint-saved copy reopens with the same slide count | yes | `065157856e816b2e…` |
| PowerPoint | `removeShape` | Remove the text box | 2 passed: text excludes 'Accent'; PowerPoint-saved copy reopens with the same slide count |  | `2b26d5c59c39c90c…` |
| PowerPoint | `insertChart` | Native column chart with embedded data | 4 passed: charts; chart data opens for editing; chart series values; PowerPoint-saved copy reopens with the same slide count | yes | `d267325d880ec162…` |
| PowerPoint | `updateChart` | Replace the chart's data | 4 passed: charts; chart data opens for editing; chart series values; PowerPoint-saved copy reopens with the same slide count | yes | `9228204559ddb65b…` |
| PowerPoint | `template` | Template chart slot bound through a template batch on the filesystem provider | 4 passed: charts; chart data opens for editing; chart series values; PowerPoint-saved copy reopens with the same slide count | yes | `b365f4e5f13c4ea9…` |
| PowerPoint | `insertImage` | Insert a picture | 2 passed: pictures; PowerPoint-saved copy reopens with the same slide count | yes | `d34686895b45d4dd…` |
| PowerPoint | `removeImage` | Remove the picture | 2 passed: pictures; PowerPoint-saved copy reopens with the same slide count |  | `f564ffce1034b0c0…` |
| PowerPoint | `backgroundImage` | Slide background picture | 2 passed: slide background is a picture; PowerPoint-saved copy reopens with the same slide count | yes | `6be4dcc0e8e739f1…` |
| PowerPoint | `insertMedia` | Embedded video with a poster frame | 2 passed: media objects; PowerPoint-saved copy reopens with the same slide count |  | `3e35f92413e418f3…` |
| PowerPoint | `insertTable` | Insert a table | 3 passed: text contains 'EMEA'; tables; PowerPoint-saved copy reopens with the same slide count | yes | `4a862903c86a507b…` |
| PowerPoint | `insertTableRows` | Append a table row | 3 passed: text contains 'APAC'; tables; PowerPoint-saved copy reopens with the same slide count |  | `dc564e3746de2f16…` |
| PowerPoint | `removeTableRows` | Remove a table row | 2 passed: tables; PowerPoint-saved copy reopens with the same slide count |  | `642337d787878f10…` |
| PowerPoint | `insertTableColumns` | Insert a table column | 3 passed: text contains 'Q2'; tables; PowerPoint-saved copy reopens with the same slide count |  | `63e39945b5e43264…` |
| PowerPoint | `removeTableColumns` | Remove a table column | 2 passed: tables; PowerPoint-saved copy reopens with the same slide count |  | `3d1a5c2ee1fa4ac4…` |
| PowerPoint | `removeTable` | Remove the table | 2 passed: tables; PowerPoint-saved copy reopens with the same slide count |  | `76890b01c86ee1d6…` |
| PowerPoint | `changeText` | Replace the title text | 2 passed: text contains 'Revised title'; PowerPoint-saved copy reopens with the same slide count |  | `eb4bff979de61a05…` |
| PowerPoint | `insert` | Insert a paragraph after the title | 2 passed: text contains 'Draft for discussion'; PowerPoint-saved copy reopens with the same slide count |  | `5155bffefbc37334…` |
| PowerPoint | `format` | Bold, coloured title | 2 passed: text contains 'Second'; PowerPoint-saved copy reopens with the same slide count | yes | `6b3dc843b7b95f86…` |
| PowerPoint | `clearStyles` | Clear direct formatting on the title | 2 passed: text contains 'Second'; PowerPoint-saved copy reopens with the same slide count |  | `769c9daddcad7463…` |
| PowerPoint | `copyStyles` | Copy paragraph formatting between lines | 2 passed: slides; PowerPoint-saved copy reopens with the same slide count |  | `769c9daddcad7463…` |
| PowerPoint | `comment` | Modern comment on a slide | 2 passed: comments at least; PowerPoint-saved copy reopens with the same slide count |  | `059b16bd97bb468e…` |
| PowerPoint | `setProperty` | Deck title property | 2 passed: title property; PowerPoint-saved copy reopens with the same slide count |  | `31880023af845c79…` |
| PowerPoint | `fill` | Fill a named template shape | 3 passed: text contains 'Northwind Traders Limited'; text excludes '[CLIENT]'; PowerPoint-saved copy reopens with the same slide count | yes | `177a699f92e828c7…` |
| Excel | `setCell` | Text, numbers and a SUM formula | 6 passed: no Excel repair log; A1 value; B2 value; B4 value; B4 formula; Excel-saved copy reopens | yes | `0d5f362a8f832fa6…` |
| Excel | `calculation` | 107 allowlisted formulas calculated by the engine, recalculated by Excel | 3 passed: no Excel repair log; formulas where Excel agrees with the cached value; Excel-saved copy reopens |  | `b811794446424d00…` |
| Excel | `setProperty` | Workbook title property | 3 passed: title property; no Excel repair log; Excel-saved copy reopens |  | `1cf68204bdeb0421…` |
| Excel | `comment` | Cell comment | 3 passed: no Excel repair log; comments at least; Excel-saved copy reopens |  | `20fe4059e2417612…` |
| Excel | `appendTableRows` | Append rows to an Excel table | 5 passed: no Excel repair log; cells contains 'APAC'; tables; table data rows grew by; Excel-saved copy reopens | yes | `4a880cfa39e3869f…` |

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
