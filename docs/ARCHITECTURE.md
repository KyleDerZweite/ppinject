# Architecture Notes

## Core design goal

Mutate only targeted PresentationML XML nodes while preserving slide structure, layout, formatting, relationships, and metadata whenever feasible.

## Current supported mutation path

1. Open the `.pptx` file as an OPC ZIP archive.
2. Read one target slide part such as `ppt/slides/slide1.xml`.
3. Read the matching slide relationships part such as `ppt/slides/_rels/slide1.xml.rels`.
4. Replace matching placeholder text inside paragraph-scoped text content.
5. Resolve picture shape names through relationship targets and replace the referenced media parts.
6. Repack the archive while leaving unrelated parts unchanged.

`highlevel.py` owns the public path and mapping adapter. `injector.py` owns archive traversal, XML mutation, media replacement and report construction. The dependency direction remains `highlevel` to `injector`; application-specific logic belongs in the caller.

The output must differ from the template. The archive owner rejects identical paths and existing file aliases before opening the output. The pre-write input checks leave an existing output untouched on failure. Write-time failures are not transactional.

## Read and write direction

- Traverse slide XML rather than rebuilding the full presentation object model.
- Target text content through explicit placeholder text replacement in paragraph content.
- Target images through existing picture shape names and slide relationship resolution.
- Preserve unsupported or unknown XML wherever possible.
- Keep relationship integrity and namespace compatibility as first-class requirements.

## Why the current text replacement works

PowerPoint often splits visible text across multiple `<a:t>` nodes inside one paragraph.
The current implementation therefore concatenates all text nodes within a paragraph,
performs placeholder replacement on that paragraph string, writes the updated text back into the first text node,
and clears the remaining paragraph text nodes.

This is good enough for deterministic template placeholders in the current MVP,
but it is not yet a full formatting-preserving run-span mutation engine.

## Preservation evidence

Synthetic tests compare every unrelated archive payload byte-for-byte, preserve part ordering and ZIP entry timestamps, compression and attributes, and compare the target slide tree after removing only the intended text difference. The fixture includes unknown namespace content, comments, geometry and formatted neighboring text. Split-run replacement retains run elements but keeps the paragraph-level formatting limitation described above. These tests do not prove byte-identical serialization of the modified slide or Windows PowerPoint opening behavior.

## Non-goals

- No full PowerPoint object model
- No slide design authoring
- No chart generation inside `ppinject`
- No generic layout editing, animation editing, or master-slide authoring
- No claims of preserving mixed run formatting inside a replaced placeholder span beyond the current paragraph-level strategy

## Integration pattern

1. Your application prepares deterministic input data.
2. Your application renders any charts or other assets separately.
3. `ppinject` replaces targeted text placeholders and image slots in the template slide.
4. Your application validates the `RenderReport` and writes workflow-specific logs.
5. Output is validated in PowerPoint on Windows as part of format-sensitive testing.

## Practical tradeoff in the current release

The current package intentionally optimizes for a narrow, reliable scenario:

- template-driven slide rendering
- known placeholder vocabulary
- known picture shape names

This keeps the API small and predictable enough for publication,
while leaving room for future expansion once more PowerPoint edge cases are covered by tests.
