# Roadmap

Current release stage: **beta, narrow MVP**.
The repository now contains a usable template-rendering path,
but the supported surface is intentionally small and still needs hardening before broader PPTX use cases.

## Phase 1 - Scaffold

- [x] Create repository metadata and collaboration docs
- [x] Create Python package skeleton
- [x] Add tests, docs, and CI structure
- [x] Capture initial architecture and roadmap notes

## Phase 2 - PPTX archive traversal MVP

- [x] Add PPTX ZIP wrapper for read and repack flow
- [x] Resolve slide relationships for media replacement in a constrained slide scenario
- [x] Add targeted placeholder replacement for paragraph-scoped text content
- [ ] Add stronger preservation tests for unchanged neighboring XML

## Phase 3 - Safe mutation primitives

- [x] Implement targeted text replacement for known template placeholders
- [x] Preserve relationship-based media replacement for existing picture shapes
- [x] Preserve archive structure for unrelated parts during write
- [ ] Validate outputs systematically against Windows PowerPoint opening behavior
- [ ] Preserve mixed run formatting inside replaced placeholder spans more precisely

## Phase 4 - API refinement

- [x] Finalize a minimal public library API for the supported scenario
- [x] Add synthetic fixture coverage for the first rendering path
- [ ] Add more fixture coverage for split runs, missing placeholders, and negative cases
- [ ] Add higher-level helper functions once the core mutation path is stable

## Phase 5 - Release readiness

- [x] Align README and docs with the actual public API
- [x] Add a publish-readiness checklist and concrete usage example
- [x] Add publish workflow automation and trusted publisher configuration
- [ ] Publish the first public package release once Windows validation is complete

## Near-term priorities

1. Add more negative and preservation-focused tests.
2. Validate generated files manually in PowerPoint on Windows for the supported template scenario.
3. Expand docs only when new behavior is actually supported.
4. Keep the supported surface narrow rather than over-claiming general PPTX mutation coverage.
