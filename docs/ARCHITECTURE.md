# Architecture Notes

## Core design goal

Mutate only targeted PresentationML XML nodes while preserving slide structure, layout, formatting, relationships, and metadata whenever feasible.

## Planned package targeting strategy

1. Open the `.pptx` file as an OPC ZIP archive.
2. Parse `ppt/presentation.xml` to locate slide ordering metadata.
3. Parse the relevant `.rels` parts to resolve slide, layout, and media targets.
4. Mutate only the intended text-bearing or tagged XML nodes in the targeted slide parts.

## Planned read and write direction

- Traverse slide XML rather than rebuilding the full presentation object model.
- Target text content through explicit shape, placeholder, or tagged-region discovery.
- Preserve unsupported or unknown XML wherever possible.
- Keep relationship integrity and namespace compatibility as first-class requirements.

## Non-goals

- No full PowerPoint object model
- No slide design authoring
- No chart, animation, or media generation in the initial phases
- No implementation in this scaffold phase

## Integration pattern

1. Your application prepares deterministic input data.
2. `ppinject` resolves the intended presentation and slide targets.
3. `ppinject` mutates only the targeted XML nodes.
4. Output is validated in PowerPoint on Windows as part of format-sensitive testing.
