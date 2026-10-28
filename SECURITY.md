# Security Policy

## Project maturity

`ppinject` is a published beta with a narrow API for targeted PPTX text and image replacement.
It parses and rewrites template archives. Automated ZIP/XML tests do not establish compatibility with every PowerPoint file or provide a sandbox for untrusted input.

## Supported versions

Security fixes are developed on `main` for the next beta release.
Older beta versions do not have maintained security branches; use the latest published beta when consuming the package.

## Reporting a vulnerability

Please report security issues privately via GitHub Security Advisories (preferred).
If that is not possible, contact maintainers directly and avoid public issue disclosure.

When reporting, include:

- affected version or commit
- reproduction steps and sample input (sanitized)
- impact assessment
- suggested fix, if available

## Response targets

- Initial acknowledgement: within 7 days
- Triage decision: within 14 days
- Fix timeline: depends on severity and maintainer availability

## Coordinated disclosure

- Please allow maintainers reasonable time to triage and patch before public disclosure.
- Remediation notes should be published alongside the relevant fix or release.

## Scope notes

- `ppinject` performs targeted PPTX XML mutations and does not intentionally execute embedded code.
- Do not process untrusted files in privileged environments without sandboxing.
- Validate outputs in your own operational context before production use.
- Changes must treat external relationships and digital signatures as safety-sensitive archive content.
- Development tooling is managed in the locked uv dev group. Runtime package dependencies are declared separately in `pyproject.toml`.
