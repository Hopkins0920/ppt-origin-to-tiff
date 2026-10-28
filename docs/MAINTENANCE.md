# Maintenance workflow / 长期维护

1. Triage incoming issues by data-loss risk, reproducibility and environment. Use `bug`, `enhancement`, `needs-info`, `compatibility` and `release-blocker` labels after the repository exists. No label automation is installed yet.
2. Confirm the current stable version and obtain a synthetic reproduction. Do not request complete confidential work directories in public.
3. Create a focused branch. Capture a failing regression for substantive bugs, make the smallest change, and run Windows CI. No runtime refactoring merely for style.
4. Changes affecting Origin/PowerPoint automation, TIFF output, geometry, cancellation or dependencies need a licensed desktop regression. Record counts, rendering checks and preserved skipped objects locally. Share only sanitized outcomes.
5. Submit a PR with validation. **The owner approves every merge.** Configure main branch protection and required CI after creation; these settings are not enforced by local files alone. No automatic merge.
6. Follow semantic versioning: patch for fixes, minor for compatible features, major for breaking input/output or workflow changes. Update engine VERSION, GUI version text, changelog, README and release notes together.
7. Before a release, follow RELEASING.md and obtain separate owner publication approval. Retain the previous release and artifacts; never silently replace or delete releases.

Review dependency update PRs weekly when practical; Dependabot is configured, but updates are never automatically merged. Check security advisories and run `pip-audit` in the proposed environment before release. The build environment snapshot is a tested baseline, not a security approval. A maintainer periodically checks new Origin/Office versions when licensed test machines are available; no unattended scheduling is activated.

Support target: latest owner-approved stable release. Compatibility rows must say tested, untested or unsupported, not assume that all versions work. Maintain a sanitized matrix in release notes; do not publish internal paths or private fixture names.
