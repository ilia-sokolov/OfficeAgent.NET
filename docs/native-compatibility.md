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

70 of 70 cases pass and 3 of 3 corrupt controls are caught, on Office 16.0.20326.20158, Microsoft Windows 11 Pro 10.0.26200. Full hashes are in the corpus manifest.

| Format | Operation | Case | Native checks | Visual | Output SHA-256 |
| --- | --- | --- | --- | --- | --- |
| Word | `changeText` | Tracked replacement of a phrase | 7 passed: revisions at least; revisions after accept all; accepted text contains 'Globex Inc.'; revisions after reject all; rejected text contains 'Acme Corp'; rejected text excludes 'Globex'; Word-saved copy reopens with the same revision count | yes | `f3b83b91c3c5e993…` |
| Word | `revision` | Two authors' tracked edits in successive commits | 9 passed: revision authors; revisions after accept all; accepted text contains 'Globex Inc.'; accepted text contains 'Northwind Traders'; revisions after reject all; rejected text contains 'Acme Corp'; rejected text excludes 'Northwind'; rejected text excludes 'Globex'; Word-saved copy reopens with the same revision count | yes | `b55fbf6a4b0eff47…` |
| Word | `insertParagraphs` | Tracked paragraph insertion | 6 passed: revisions at least; revisions after accept all; accepted text contains 'Inserted obligation paragraph.'; revisions after reject all; rejected text excludes 'Inserted obligation paragraph.'; Word-saved copy reopens with the same revision count |  | `a27cfb11d75b0b5d…` |
| Word | `removeParagraph` | Tracked paragraph removal | 6 passed: revisions at least; revisions after accept all; accepted text excludes 'Service Agreement'; revisions after reject all; rejected text contains 'Service Agreement'; Word-saved copy reopens with the same revision count |  | `9e682a50dc6ef156…` |
| Word | `format` | Tracked formatting change | 4 passed: revisions at least; revisions after accept all; accepted text contains 'Acme Corp'; Word-saved copy reopens with the same revision count |  | `70ef5f087f4ae89a…` |
| Word | `insert` | Direct paragraph insertion after a found phrase | 3 passed: revisions; text contains 'The Supplier is Acme Corp.'; Word-saved copy reopens with the same revision count |  | `58e3f06654047f1e…` |
| Word | `revision` | Accept every pending revision | 4 passed: revisions; text contains 'approved'; text excludes 'draft'; Word-saved copy reopens with the same revision count |  | `226c10c440c4de55…` |
| Word | `revision` | Reject every pending revision | 4 passed: revisions; text contains 'draft'; text excludes 'approved'; Word-saved copy reopens with the same revision count |  | `d1120bbc067a3a0d…` |
| Word | `comment` | Comment beside existing review threads | 3 passed: text contains 'Numbered obligation'; comments at least; Word-saved copy reopens with the same revision count | yes | `991a11c530146b5f…` |
| Word | `note` | Footnote | 3 passed: footnotes; footnote text; Word-saved copy reopens with the same revision count | yes | `4bdb104a9404f903…` |
| Word | `note` | Endnote | 3 passed: endnotes; endnote text; Word-saved copy reopens with the same revision count |  | `c9a4af41ef3f95d9…` |
| Word | `fill` | Fill a nested content control | 3 passed: text contains 'Contoso Holdings'; content controls at least; Word-saved copy reopens with the same revision count |  | `47cf36bb2fcc9094…` |
| Word | `insertTable` | Insert a 3x3 table | 5 passed: text contains 'Widget'; tables; table rows; table columns; Word-saved copy reopens with the same revision count | yes | `992ae785ea2c73ad…` |
| Word | `insertTableRows` | Append table rows | 4 passed: text contains 'Sprocket'; tables; table rows; Word-saved copy reopens with the same revision count |  | `6d68b5c4ba721e1c…` |
| Word | `removeTableRows` | Remove a table row | 4 passed: text excludes 'Gadget'; tables; table rows; Word-saved copy reopens with the same revision count |  | `0a975a57459c8765…` |
| Word | `insertTableColumns` | Insert a table column | 4 passed: text contains 'Total'; tables; table columns; Word-saved copy reopens with the same revision count |  | `a7365cb308deb8fa…` |
| Word | `removeTableColumns` | Remove a table column | 4 passed: text excludes 'Price'; tables; table columns; Word-saved copy reopens with the same revision count |  | `2fe97c1aca628634…` |
| Word | `removeTableColumns` | Tracked column removal, resolved in Word | 9 passed: revisions at least; table columns; revisions after accept all; accepted text excludes 'Price'; table columns after accept all; revisions after reject all; rejected text contains 'Price'; table columns after reject all; Word-saved copy reopens with the same revision count |  | `4e32678f9116992b…` |
| Word | `repeatTableRow` | Repeat a template row per record | 7 passed: text contains 'Widget'; text contains 'Gadget'; text contains 'Sprocket'; text excludes '{{'; tables; table rows; Word-saved copy reopens with the same revision count | yes | `2e1cefb2ae2750bc…` |
| Word | `removeTable` | Remove a table | 3 passed: text excludes 'Widget'; tables; Word-saved copy reopens with the same revision count |  | `51343ce28e71bc27…` |
| Word | `insertImage` | Insert an inline picture | 2 passed: inline pictures; Word-saved copy reopens with the same revision count | yes | `763be6294574a7b8…` |
| Word | `removeImage` | Remove the picture | 2 passed: inline pictures; Word-saved copy reopens with the same revision count |  | `097ea5dd277e8e2b…` |
| Word | `backgroundImage` | Page background picture | 2 passed: header holds a picture behind text; Word-saved copy reopens with the same revision count | yes | `752c5d7d49c55e95…` |
| Word | `headerFooter` | Header, footer and page number | 3 passed: header contains 'Contoso'; footer contains 'Services agreement'; Word-saved copy reopens with the same revision count | yes | `bdb0d30466e00af2…` |
| Word | `pageSetup` | Landscape A4 with wide margins | 2 passed: landscape; Word-saved copy reopens with the same revision count | yes | `3c96538f8996d022…` |
| Word | `insertBreak` | Page break | 2 passed: pages at least; Word-saved copy reopens with the same revision count | yes | `8026e294a3036e8d…` |
| Word | `setProperty` | Document title property | 2 passed: title property; Word-saved copy reopens with the same revision count |  | `f7586a7e73c004b2…` |
| Word | `defineStyle` | Define and apply a heading style | 2 passed: style 'Contract Heading' exists; Word-saved copy reopens with the same revision count |  | `342d0fd596b04532…` |
| Word | `clearStyles` | Clear direct formatting on a bolded phrase | 3 passed: revisions; Word reads the same text as in the input; Word-saved copy reopens with the same revision count |  | `51343ce28e71bc27…` |
| Word | `clearStyles` | Clear formatting on runs written without xml:space | 2 passed: text contains 'Quarterly target'; Word-saved copy reopens with the same revision count |  | `14083271bed742cd…` |
| Word | `copyStyles` | Copy direct formatting between paragraphs | 2 passed: Word reads the same text as in the input; Word-saved copy reopens with the same revision count |  | `c2713a3eeca90607…` |
| Word | `create` | Blank document from OfficeAgent, then written | 4 passed: revisions; text contains 'Created by OfficeAgent.'; Word compatibility mode; Word-saved copy reopens with the same revision count |  | `0ab30d361f27d81d…` |
| Word | `template` | Template with a slot and a repeating row, populated | 7 passed: text contains 'Fabrikam Ltd'; text contains 'Consulting'; text contains 'Support'; text excludes '{{'; text excludes 'CUSTOMER'; table rows; Word-saved copy reopens with the same revision count | yes | `1fd8f00c820990de…` |
| Word | `template` | Template image slot bound through a template batch on the filesystem provider | 5 passed: text contains 'Quote for Fabrikam Ltd'; text excludes 'CUSTOMER'; picture size in points; inline pictures; Word-saved copy reopens with the same revision count | yes | `0d962804e144d7fc…` |
| Word | `comparison` | Comparison redline of a table-cell edit | 10 passed: revisions at least; tables; table rows; revisions after accept all; accepted text contains 'Gizmo'; accepted text excludes 'Gadget'; revisions after reject all; rejected text contains 'Gadget'; rejected text excludes 'Gizmo'; Word-saved copy reopens with the same revision count | yes | `c2f25f0d563d1855…` |
| Word | `comparison` | Comparison redline between two versions | 7 passed: revisions at least; revisions after accept all; accepted text contains 'Globex Inc.'; revisions after reject all; rejected text contains 'Acme Corp'; rejected text excludes 'Globex'; Word-saved copy reopens with the same revision count | yes | `f7e09176b1aef46d…` |
| Word | `assembly` | Two documents assembled, each keeping its header | 6 passed: revisions; text contains 'Acme Corp'; text contains 'Appendix A: rate card.'; sections at least; Word compatibility mode; Word-saved copy reopens with the same revision count | yes | `7138b9aea55e5044…` |
| PowerPoint | `insertSlide` | Two slides added to a blank deck | 4 passed: slides; text contains 'Second'; text contains 'Closing point'; PowerPoint-saved copy reopens with the same slide count | yes | `dfe28ff5f346653a…` |
| PowerPoint | `removeSlide` | Remove a slide | 3 passed: slides; text excludes 'Another point'; PowerPoint-saved copy reopens with the same slide count |  | `48bef843a8dfb9ea…` |
| PowerPoint | `moveSlide` | Move the last slide first | 3 passed: slides; first slide contains 'Third'; PowerPoint-saved copy reopens with the same slide count |  | `d6cae28a7284f5ba…` |
| PowerPoint | `duplicateSlide` | Duplicate a slide | 2 passed: slides; PowerPoint-saved copy reopens with the same slide count |  | `c97dd2baf9177d98…` |
| PowerPoint | `section` | Add a section | 3 passed: sections at least; section 'Financials'; PowerPoint-saved copy reopens with the same slide count |  | `67b46cad33bc8586…` |
| PowerPoint | `transition` | Fade transition on every slide | 3 passed: slides; every slide has a transition; PowerPoint-saved copy reopens with the same slide count |  | `2a8508a6c5e8b5ee…` |
| PowerPoint | `animate` | Fade-in animation on a shape | 2 passed: animations at least; PowerPoint-saved copy reopens with the same slide count |  | `8ddf0ff5dd7c9730…` |
| PowerPoint | `headerFooter` | Footer and slide numbers | 2 passed: footer contains 'Confidential'; PowerPoint-saved copy reopens with the same slide count | yes | `984178b840e00c2d…` |
| PowerPoint | `insertShape` | Insert a text box | 2 passed: text contains 'Accent'; PowerPoint-saved copy reopens with the same slide count | yes | `602c99f681e1aa45…` |
| PowerPoint | `removeShape` | Remove the text box | 2 passed: text excludes 'Accent'; PowerPoint-saved copy reopens with the same slide count |  | `4c1fdeb8138714a1…` |
| PowerPoint | `insertChart` | Native column chart with embedded data | 4 passed: charts; chart data opens for editing; chart series values; PowerPoint-saved copy reopens with the same slide count | yes | `8c3a0a2f4ea4dfb0…` |
| PowerPoint | `updateChart` | Replace the chart's data | 4 passed: charts; chart data opens for editing; chart series values; PowerPoint-saved copy reopens with the same slide count | yes | `8c963a91c2432a44…` |
| PowerPoint | `template` | Template chart slot bound through a template batch on the filesystem provider | 4 passed: charts; chart data opens for editing; chart series values; PowerPoint-saved copy reopens with the same slide count | yes | `76c1246aaf5714e2…` |
| PowerPoint | `insertImage` | Insert a picture | 2 passed: pictures; PowerPoint-saved copy reopens with the same slide count | yes | `289a22f0d2070cbb…` |
| PowerPoint | `removeImage` | Remove the picture | 2 passed: pictures; PowerPoint-saved copy reopens with the same slide count |  | `1bc96bf2843d4ba2…` |
| PowerPoint | `backgroundImage` | Slide background picture | 2 passed: slide background is a picture; PowerPoint-saved copy reopens with the same slide count | yes | `2f9a36237391f7a3…` |
| PowerPoint | `insertMedia` | Embedded video with a poster frame | 2 passed: media objects; PowerPoint-saved copy reopens with the same slide count |  | `ce97e9091f951d68…` |
| PowerPoint | `insertTable` | Insert a table | 3 passed: text contains 'EMEA'; tables; PowerPoint-saved copy reopens with the same slide count | yes | `5f0b2822c3d230f7…` |
| PowerPoint | `insertTableRows` | Append a table row | 3 passed: text contains 'APAC'; tables; PowerPoint-saved copy reopens with the same slide count |  | `0a028432de2eee8b…` |
| PowerPoint | `removeTableRows` | Remove a table row | 2 passed: tables; PowerPoint-saved copy reopens with the same slide count |  | `7ad08242bccf204a…` |
| PowerPoint | `insertTableColumns` | Insert a table column | 3 passed: text contains 'Q2'; tables; PowerPoint-saved copy reopens with the same slide count |  | `22fe3dddc7b014a9…` |
| PowerPoint | `removeTableColumns` | Remove a table column | 2 passed: tables; PowerPoint-saved copy reopens with the same slide count |  | `dca62908274ea9f8…` |
| PowerPoint | `removeTable` | Remove the table | 2 passed: tables; PowerPoint-saved copy reopens with the same slide count |  | `c4feff922463c508…` |
| PowerPoint | `changeText` | Replace the title text | 2 passed: text contains 'Revised title'; PowerPoint-saved copy reopens with the same slide count |  | `e41f9ccd2bfb3287…` |
| PowerPoint | `insert` | Insert a paragraph after the title | 2 passed: text contains 'Draft for discussion'; PowerPoint-saved copy reopens with the same slide count |  | `120ffeb486ade78e…` |
| PowerPoint | `format` | Bold, coloured title | 2 passed: text contains 'Second'; PowerPoint-saved copy reopens with the same slide count | yes | `45e567892fc2e294…` |
| PowerPoint | `clearStyles` | Clear direct formatting on the title | 2 passed: text contains 'Second'; PowerPoint-saved copy reopens with the same slide count |  | `7a5721bbf3fe598d…` |
| PowerPoint | `copyStyles` | Copy paragraph formatting between lines | 2 passed: slides; PowerPoint-saved copy reopens with the same slide count |  | `7a5721bbf3fe598d…` |
| PowerPoint | `comment` | Modern comment on a slide | 2 passed: comments at least; PowerPoint-saved copy reopens with the same slide count |  | `bed901e681c325c4…` |
| PowerPoint | `fill` | Fill a named template shape | 3 passed: text contains 'Northwind Traders Limited'; text excludes '[CLIENT]'; PowerPoint-saved copy reopens with the same slide count | yes | `c6ab8e36cce020f4…` |
| Excel | `setCell` | Text, numbers and a SUM formula | 6 passed: no Excel repair log; A1 value; B2 value; B4 value; B4 formula; Excel-saved copy reopens | yes | `42a4f306341527d6…` |
| Excel | `comment` | Cell comment | 3 passed: no Excel repair log; comments at least; Excel-saved copy reopens |  | `311bbfcf9ac9a9c0…` |
| Excel | `appendTableRows` | Append rows to an Excel table | 5 passed: no Excel repair log; cells contains 'APAC'; tables; table data rows grew by; Excel-saved copy reopens | yes | `d37069b09137a4ff…` |

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
