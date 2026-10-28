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

- A clean portable build using `requirements-portable-conda.txt`, `requirements-build.txt` and `scripts/build_portable.py`.
- Review of generated LICENSES, Python/Tcl/Tk notices and the full native DLL inventory.
- Confirmation that the bundle contains no Origin/Office binaries, activation material, user data or unlicensed third-party content.
- Smoke tests and native conversion on an independent Windows computer.
- Release notes stating the tested environment and signing status.
- Owner approval of the exact assets, followed by archive/checksum verification.

The Windows portable checks workflow builds and checks the preview on an independent Windows runner. It does not upload binaries or create releases. Native export and manual desktop startup results are tracked separately in [portable validation](PORTABLE_VALIDATION.md). The older `scripts/build_windows.py` remains a local candidate builder; it is not used for portable distribution.

To reproduce the portable preview, create a Windows conda prefix from `requirements-portable-conda.txt`, install `requirements-build.txt` using that prefix's `python -I -m pip --isolated`, then run `python -I scripts/build_portable.py`. Retain conda package caches until the notice/recipe collection completes. Run `scripts/check_portable.py` against the resulting ZIP. Each build uses a new directory and keeps previous builds intact.

## Maintenance and rollback

Keep prior releases and artifacts. Fixes use reviewed branches and new versions. Revert faulty changes through a pull request; do not silently replace published assets, rewrite public history or delete earlier releases.
