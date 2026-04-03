# Security Policy

## Project maturity

`ppinject` is in an **early scaffold stage**.
The repository structure is in place, but the PPTX processing implementation is not yet complete.

## Supported versions

At this stage, only the latest `main` branch state is considered supported for security fixes.

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

- `ppinject` is expected to perform targeted PPTX XML mutations and should not execute embedded code.
- Do not process untrusted files in privileged environments without sandboxing.
- Validate outputs in your own operational context before production use.
- Future implementations must treat external relationships and digital signatures as safety-sensitive surfaces.
