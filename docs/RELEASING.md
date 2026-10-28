# Release checklist (manual approval gates)

- [ ] Owner approves repository name, account, visibility, license and attribution. Add the approved LICENSE; remove pending-publication notices only after approval.
- [ ] Inspect all tracked files and run `python scripts/audit_source.py`; audit the complete Git history before first upload. `.gitignore` is not a history scrubber.
- [ ] Unit tests and GUI smoke checks pass; Windows 3.11/3.12 CI results reviewed after upload. Do not claim CI has run before GitHub executes it.
- [ ] Run `python -m pip_audit -r requirements.txt` in an isolated environment; review pinned build dependencies and runtime advisories as well. Resolve or explicitly document any release blockers.
- [ ] Native local test: strict success; skip failure then successful later graph; all skipped gives no output; cancellation publishes no result; original checksum unchanged; geometry/groups and non-target objects preserved; PowerPoint opens/renders the candidate. Use private/synthetic fixtures, not committed manuscripts.
- [ ] Build with `python scripts/build_windows.py` in a clean Windows CPython environment with the approved dependency snapshot. Do not copy old `_internal` folders into a new release.
- [ ] Inspect generated LICENSES and native DLL inventory; resolve transitive license notices. Confirm no Origin/Office binaries, fonts, license material or user data are bundled.
- [ ] Verify EXE smoke check, independent machine startup and native export. Candidate is unsigned unless code signing is explicitly arranged; never weaken Windows security policy.
- [ ] Create source ZIP from the approved Git commit (`git archive`), Windows folder ZIP, SHA256SUMS and release notes with exact environment and known limitations. Do not add unpublished research screenshots as examples.
- [ ] Owner reviews diff and approves merging the preparation branch to main. Owner separately approves upload/publication. No automatic GitHub release publishing workflow is configured.
- [ ] After upload, enable branch protection, required CI checks and private vulnerability reporting. Create an owner-approved GitHub release, attach reviewed assets, verify hashes and links.

The supplied Windows workflow creates temporary Actions artifacts only when manually dispatched. Artifact access and retention follow GitHub permissions; treat a public repository's build artifacts as public distribution too. Do not dispatch until license and binary review approval is in place.

Rollback: keep the previous stable release and revert through a reviewed PR. Do not rewrite public history or delete existing files/releases without explicit owner permission.
