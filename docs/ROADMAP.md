# Roadmap

Current release stage: **pre-implementation scaffold**.
The repository layout and tooling are in place, but the injector behavior is still to be built.

## Phase 1 - Scaffold

- [x] Create repository metadata and collaboration docs
- [x] Create Python package skeleton
- [x] Add tests, docs, and CI structure
- [x] Capture initial architecture and roadmap notes

## Phase 2 - PPTX archive traversal MVP

- [ ] Add PPTX ZIP wrapper for read and repack flow
- [ ] Resolve presentation -> slide targets through PresentationML relationships
- [ ] Add targeted text node discovery for a constrained slide scenario
- [ ] Add preservation tests for unchanged neighboring XML

## Phase 3 - Safe mutation primitives

- [ ] Implement targeted text replacement for selected shapes or placeholders
- [ ] Escape XML-sensitive characters safely
- [ ] Preserve relationships, namespaces, and ordering required by PowerPoint
- [ ] Validate outputs against Windows PowerPoint opening behavior

## Phase 4 - API refinement

- [ ] Finalize a minimal public library API
- [ ] Add synthetic fixture coverage for common slide structures
- [ ] Add higher-level helper functions once the core mutation path is stable

## Phase 5 - Release readiness

- [ ] Add publish workflow and release process docs
- [ ] Publish first public package release when behavior is usable
