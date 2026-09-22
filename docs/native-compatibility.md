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
Each case runs in its own process with a time limit, so a modal dialog becomes a recorded
failure rather than a stalled run.

The harness was also shown to fail on wrong output. Each tampered copy below was given a
matching manifest hash, so only the checks could catch it:

| Tamper | Result |
| --- | --- |
| Tracked `changeText` output replaced by its input | Fails: 0 revisions, accepted text lacks the replacement |
| Second author of a two-author redline renamed to the first | Fails: revision authors differ |
| `SUM(B2:B3)` rewritten to `SUM(B2:B2)` | Fails: Excel's recalculated value and formula differ |
| Chart's link to its embedded workbook removed | Fails: PowerPoint does not answer (killed at the time limit) |
| Chart's embedded workbook replaced by junk | Fails: PowerPoint refuses to open the deck |

## Matrix

<!-- native-matrix:begin -->

70 of 70 cases pass and 3 of 3 corrupt controls are caught, on Office 16.0.20326.20158, Microsoft Windows 11 Pro 10.0.26200. Full hashes are in the corpus manifest.

| Format | Operation | Case | Native checks | Visual | Output SHA-256 |
| --- | --- | --- | --- | --- | --- |
| Word | `changeText` | Tracked replacement of a phrase | 7 passed: revisions at least; revisions after accept all; accepted text contains 'Globex Inc.'; revisions after reject all; rejected text contains 'Acme Corp'; rejected text excludes 'Globex'; Word-saved copy reopens with the same revision count | yes | `92f691bd9674b5d0…` |
| Word | `revision` | Two authors' tracked edits in successive commits | 9 passed: revision authors; revisions after accept all; accepted text contains 'Globex Inc.'; accepted text contains 'Northwind Traders'; revisions after reject all; rejected text contains 'Acme Corp'; rejected text excludes 'Northwind'; rejected text excludes 'Globex'; Word-saved copy reopens with the same revision count | yes | `5ed52f35f730edb1…` |
| Word | `insertParagraphs` | Tracked paragraph insertion | 6 passed: revisions at least; revisions after accept all; accepted text contains 'Inserted obligation paragraph.'; revisions after reject all; rejected text excludes 'Inserted obligation paragraph.'; Word-saved copy reopens with the same revision count |  | `fba501c7fed86082…` |
| Word | `removeParagraph` | Tracked paragraph removal | 6 passed: revisions at least; revisions after accept all; accepted text excludes 'Service Agreement'; revisions after reject all; rejected text contains 'Service Agreement'; Word-saved copy reopens with the same revision count |  | `dfd51d216ae94ec8…` |
| Word | `format` | Tracked formatting change | 4 passed: revisions at least; revisions after accept all; accepted text contains 'Acme Corp'; Word-saved copy reopens with the same revision count |  | `9b6808c87f38fc7c…` |
| Word | `insert` | Direct paragraph insertion after a found phrase | 3 passed: revisions; text contains 'The Supplier is Acme Corp.'; Word-saved copy reopens with the same revision count |  | `8de79ed994969b79…` |
| Word | `revision` | Accept every pending revision | 4 passed: revisions; text contains 'approved'; text excludes 'draft'; Word-saved copy reopens with the same revision count |  | `226c10c440c4de55…` |
| Word | `revision` | Reject every pending revision | 4 passed: revisions; text contains 'draft'; text excludes 'approved'; Word-saved copy reopens with the same revision count |  | `d1120bbc067a3a0d…` |
| Word | `comment` | Comment beside existing review threads | 3 passed: text contains 'Numbered obligation'; comments at least; Word-saved copy reopens with the same revision count | yes | `0eb7d64c534d7828…` |
| Word | `note` | Footnote | 3 passed: footnotes; footnote text; Word-saved copy reopens with the same revision count | yes | `a7baafaae8adc17e…` |
| Word | `note` | Endnote | 3 passed: endnotes; endnote text; Word-saved copy reopens with the same revision count |  | `8d7f8391a92412d9…` |
| Word | `fill` | Fill a nested content control | 3 passed: text contains 'Contoso Holdings'; content controls at least; Word-saved copy reopens with the same revision count |  | `47cf36bb2fcc9094…` |
| Word | `insertTable` | Insert a 3x3 table | 5 passed: text contains 'Widget'; tables; table rows; table columns; Word-saved copy reopens with the same revision count | yes | `08a1513cec2b98bd…` |
| Word | `insertTableRows` | Append table rows | 4 passed: text contains 'Sprocket'; tables; table rows; Word-saved copy reopens with the same revision count |  | `983818cf619ca9fe…` |
| Word | `removeTableRows` | Remove a table row | 4 passed: text excludes 'Gadget'; tables; table rows; Word-saved copy reopens with the same revision count |  | `1ead2ca3bff1344a…` |
| Word | `insertTableColumns` | Insert a table column | 4 passed: text contains 'Total'; tables; table columns; Word-saved copy reopens with the same revision count |  | `7dddfdb3d4c00be3…` |
| Word | `removeTableColumns` | Remove a table column | 4 passed: text excludes 'Price'; tables; table columns; Word-saved copy reopens with the same revision count |  | `69eb3d747449ab4d…` |
| Word | `removeTableColumns` | Tracked column removal, resolved in Word | 9 passed: revisions at least; table columns; revisions after accept all; accepted text excludes 'Price'; table columns after accept all; revisions after reject all; rejected text contains 'Price'; table columns after reject all; Word-saved copy reopens with the same revision count |  | `58be6233c9ace10d…` |
| Word | `repeatTableRow` | Repeat a template row per record | 7 passed: text contains 'Widget'; text contains 'Gadget'; text contains 'Sprocket'; text excludes '{{'; tables; table rows; Word-saved copy reopens with the same revision count | yes | `a820c8d340cb67db…` |
| Word | `removeTable` | Remove a table | 3 passed: text excludes 'Widget'; tables; Word-saved copy reopens with the same revision count |  | `c50ec2d370000a3e…` |
| Word | `insertImage` | Insert an inline picture | 2 passed: inline pictures; Word-saved copy reopens with the same revision count | yes | `68b132f8aff2e7fe…` |
| Word | `removeImage` | Remove the picture | 2 passed: inline pictures; Word-saved copy reopens with the same revision count |  | `af9228dcc04dba51…` |
| Word | `backgroundImage` | Page background picture | 2 passed: header holds a picture behind text; Word-saved copy reopens with the same revision count | yes | `8edcbb99e89079d0…` |
| Word | `headerFooter` | Header, footer and page number | 3 passed: header contains 'Contoso'; footer contains 'Services agreement'; Word-saved copy reopens with the same revision count | yes | `d688f5e944e28225…` |
| Word | `pageSetup` | Landscape A4 with wide margins | 2 passed: landscape; Word-saved copy reopens with the same revision count | yes | `b1e0012817429119…` |
| Word | `insertBreak` | Page break | 2 passed: pages at least; Word-saved copy reopens with the same revision count | yes | `86335d964d229a33…` |
| Word | `setProperty` | Document title property | 2 passed: title property; Word-saved copy reopens with the same revision count |  | `2b422ee18aebc464…` |
| Word | `defineStyle` | Define and apply a heading style | 2 passed: style 'Contract Heading' exists; Word-saved copy reopens with the same revision count |  | `d24deb1fc6303ece…` |
| Word | `clearStyles` | Clear direct formatting on a bolded phrase | 3 passed: revisions; Word reads the same text as in the input; Word-saved copy reopens with the same revision count |  | `c50ec2d370000a3e…` |
| Word | `clearStyles` | Clear formatting on runs written without xml:space | 2 passed: text contains 'Quarterly target'; Word-saved copy reopens with the same revision count |  | `14083271bed742cd…` |
| Word | `copyStyles` | Copy direct formatting between paragraphs | 2 passed: Word reads the same text as in the input; Word-saved copy reopens with the same revision count |  | `c2713a3eeca90607…` |
| Word | `create` | Blank document from OfficeAgent, then written | 4 passed: revisions; text contains 'Created by OfficeAgent.'; Word compatibility mode; Word-saved copy reopens with the same revision count |  | `a7085384a5c7b4ed…` |
| Word | `template` | Template with a slot and a repeating row, populated | 7 passed: text contains 'Fabrikam Ltd'; text contains 'Consulting'; text contains 'Support'; text excludes '{{'; text excludes 'CUSTOMER'; table rows; Word-saved copy reopens with the same revision count | yes | `f9c70a3cdd9747df…` |
| Word | `template` | Template image slot bound through a template batch on the filesystem provider | 5 passed: text contains 'Quote for Fabrikam Ltd'; text excludes 'CUSTOMER'; picture size in points; inline pictures; Word-saved copy reopens with the same revision count | yes | `bba938e62b746af9…` |
| Word | `comparison` | Comparison redline of a table-cell edit | 10 passed: revisions at least; tables; table rows; revisions after accept all; accepted text contains 'Gizmo'; accepted text excludes 'Gadget'; revisions after reject all; rejected text contains 'Gadget'; rejected text excludes 'Gizmo'; Word-saved copy reopens with the same revision count | yes | `7c633d16c9e517b4…` |
| Word | `comparison` | Comparison redline between two versions | 7 passed: revisions at least; revisions after accept all; accepted text contains 'Globex Inc.'; revisions after reject all; rejected text contains 'Acme Corp'; rejected text excludes 'Globex'; Word-saved copy reopens with the same revision count | yes | `752cd35f37b9b4d8…` |
| Word | `assembly` | Two documents assembled, each keeping its header | 6 passed: revisions; text contains 'Acme Corp'; text contains 'Appendix A: rate card.'; sections at least; Word compatibility mode; Word-saved copy reopens with the same revision count | yes | `3f3fee4ec5d1ec75…` |
| PowerPoint | `insertSlide` | Two slides added to a blank deck | 4 passed: slides; text contains 'Second'; text contains 'Closing point'; PowerPoint-saved copy reopens with the same slide count | yes | `36093d6ca9d30894…` |
| PowerPoint | `removeSlide` | Remove a slide | 3 passed: slides; text excludes 'Another point'; PowerPoint-saved copy reopens with the same slide count |  | `ef1f6bc707964ae5…` |
| PowerPoint | `moveSlide` | Move the last slide first | 3 passed: slides; first slide contains 'Third'; PowerPoint-saved copy reopens with the same slide count |  | `256f166e1f02194c…` |
| PowerPoint | `duplicateSlide` | Duplicate a slide | 2 passed: slides; PowerPoint-saved copy reopens with the same slide count |  | `0358aa598d1971bd…` |
| PowerPoint | `section` | Add a section | 3 passed: sections at least; section 'Financials'; PowerPoint-saved copy reopens with the same slide count |  | `bcb0c61870306cdb…` |
| PowerPoint | `transition` | Fade transition on every slide | 3 passed: slides; every slide has a transition; PowerPoint-saved copy reopens with the same slide count |  | `8313f2e903e3c598…` |
| PowerPoint | `animate` | Fade-in animation on a shape | 2 passed: animations at least; PowerPoint-saved copy reopens with the same slide count |  | `35a2e7ba9f6f311e…` |
| PowerPoint | `headerFooter` | Footer and slide numbers | 2 passed: footer contains 'Confidential'; PowerPoint-saved copy reopens with the same slide count | yes | `8bdb9c06d4d14487…` |
| PowerPoint | `insertShape` | Insert a text box | 2 passed: text contains 'Accent'; PowerPoint-saved copy reopens with the same slide count | yes | `0baba9f73cc93433…` |
| PowerPoint | `removeShape` | Remove the text box | 2 passed: text excludes 'Accent'; PowerPoint-saved copy reopens with the same slide count |  | `230d2d364937d912…` |
| PowerPoint | `insertChart` | Native column chart with embedded data | 4 passed: charts; chart data opens for editing; chart series values; PowerPoint-saved copy reopens with the same slide count | yes | `ed2cfb4bbea98c7d…` |
| PowerPoint | `updateChart` | Replace the chart's data | 4 passed: charts; chart data opens for editing; chart series values; PowerPoint-saved copy reopens with the same slide count | yes | `df3c09d612d701e9…` |
| PowerPoint | `template` | Template chart slot bound through a template batch on the filesystem provider | 4 passed: charts; chart data opens for editing; chart series values; PowerPoint-saved copy reopens with the same slide count | yes | `1976b1880a3ba4e3…` |
| PowerPoint | `insertImage` | Insert a picture | 2 passed: pictures; PowerPoint-saved copy reopens with the same slide count | yes | `b76187d3797dc265…` |
| PowerPoint | `removeImage` | Remove the picture | 2 passed: pictures; PowerPoint-saved copy reopens with the same slide count |  | `372a6d2a561737f4…` |
| PowerPoint | `backgroundImage` | Slide background picture | 2 passed: slide background is a picture; PowerPoint-saved copy reopens with the same slide count | yes | `21604a2721d6f3f4…` |
| PowerPoint | `insertMedia` | Embedded video with a poster frame | 2 passed: media objects; PowerPoint-saved copy reopens with the same slide count |  | `2e3875978afd3326…` |
| PowerPoint | `insertTable` | Insert a table | 3 passed: text contains 'EMEA'; tables; PowerPoint-saved copy reopens with the same slide count | yes | `584ec8d98569ae52…` |
| PowerPoint | `insertTableRows` | Append a table row | 3 passed: text contains 'APAC'; tables; PowerPoint-saved copy reopens with the same slide count |  | `3b417aceb358d26a…` |
| PowerPoint | `removeTableRows` | Remove a table row | 2 passed: tables; PowerPoint-saved copy reopens with the same slide count |  | `b48e13fc7b1df3fe…` |
| PowerPoint | `insertTableColumns` | Insert a table column | 3 passed: text contains 'Q2'; tables; PowerPoint-saved copy reopens with the same slide count |  | `52fcae37a82a3322…` |
| PowerPoint | `removeTableColumns` | Remove a table column | 2 passed: tables; PowerPoint-saved copy reopens with the same slide count |  | `da67ae1e0e7db106…` |
| PowerPoint | `removeTable` | Remove the table | 2 passed: tables; PowerPoint-saved copy reopens with the same slide count |  | `7e447b8503efe388…` |
| PowerPoint | `changeText` | Replace the title text | 2 passed: text contains 'Revised title'; PowerPoint-saved copy reopens with the same slide count |  | `19d81a0678b4fa33…` |
| PowerPoint | `insert` | Insert a paragraph after the title | 2 passed: text contains 'Draft for discussion'; PowerPoint-saved copy reopens with the same slide count |  | `4b76ce3c282fa499…` |
| PowerPoint | `format` | Bold, coloured title | 2 passed: text contains 'Second'; PowerPoint-saved copy reopens with the same slide count | yes | `fcfcd47ae791511d…` |
| PowerPoint | `clearStyles` | Clear direct formatting on the title | 2 passed: text contains 'Second'; PowerPoint-saved copy reopens with the same slide count |  | `fe7cbd727687163d…` |
| PowerPoint | `copyStyles` | Copy paragraph formatting between lines | 2 passed: slides; PowerPoint-saved copy reopens with the same slide count |  | `fe7cbd727687163d…` |
| PowerPoint | `comment` | Modern comment on a slide | 2 passed: comments at least; PowerPoint-saved copy reopens with the same slide count |  | `9d9374dd55761f0a…` |
| PowerPoint | `fill` | Fill a named template shape | 3 passed: text contains 'Northwind Traders Limited'; text excludes '[CLIENT]'; PowerPoint-saved copy reopens with the same slide count | yes | `57e395213a71a84c…` |
| Excel | `setCell` | Text, numbers and a SUM formula | 6 passed: no Excel repair log; A1 value; B2 value; B4 value; B4 formula; Excel-saved copy reopens | yes | `86a397bf84e349ef…` |
| Excel | `comment` | Cell comment | 3 passed: no Excel repair log; comments at least; Excel-saved copy reopens |  | `b04ac5b9c6cc8519…` |
| Excel | `appendTableRows` | Append rows to an Excel table | 5 passed: no Excel repair log; cells contains 'APAC'; tables; table data rows grew by; Excel-saved copy reopens | yes | `b6c96755328e621f…` |

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
