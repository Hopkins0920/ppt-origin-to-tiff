# Releasing

## Source releases

1. Create a release branch from `main`. Update the engine/GUI version when behavior changes, the changelog and a versioned release note.
2. Update `.github/release-source.toml` with the intended tag and release note path. This file is the source publication trigger.
3. Review the full diff and tracked files. Run the unit tests, GUI smoke check, source audit and dependency advisory checks. Changes affecting native export, geometry, cancellation or dependencies also require a licensed desktop regression.
4. Open a pull request. The owner reviews the change and approves merging and publication. Merging a change to the release configuration starts the source release workflow.
5. The workflow runs Windows checks on Python 3.11 and 3.12, audits build dependencies, then creates the source ZIP, SHA256SUMS, tag and GitHub release for that commit. It does not overwrite an existing tag or release. A workflow failure leaves publication incomplete; inspect Actions before retrying.
6. Verify the release page, assets, checksums, version and installation instructions. Repository visibility is a separate owner-controlled setting.

Source archives contain tracked files from the release commit. Real presentations, native fixtures, reports, virtual environments and executable builds are excluded. Audit Git history before making a repository public; ignore rules do not remove existing history.

## Windows executable releases

Source publication does not publish a Windows executable. An executable release additionally requires:

- A clean build using `requirements-build.txt` and `scripts/build_windows.py`.
- Review of generated LICENSES, Python/Tcl/Tk notices and the full native DLL inventory.
- Confirmation that the bundle contains no Origin/Office binaries, activation material, user data or unlicensed third-party content.
- Smoke tests and native conversion on an independent Windows computer.
- Release notes stating the tested environment and signing status.
- Owner approval of the exact assets, followed by archive/checksum verification.

The manual Windows build workflow produces temporary Actions artifacts. In a public repository those artifacts may be publicly downloadable, so it must only be dispatched after bundle distribution checks are complete.

## Maintenance and rollback

Keep prior releases and artifacts. Fixes use reviewed branches and new versions. Revert faulty changes through a pull request; do not silently replace published assets, rewrite public history or delete earlier releases.
