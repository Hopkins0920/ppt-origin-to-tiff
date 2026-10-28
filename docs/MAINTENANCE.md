# Maintenance workflow / 长期维护

1. Triage issues by reproducibility, data-loss risk and environment. Useful labels include `bug`, `enhancement`, `needs-info`, `compatibility` and `release-blocker`.
2. Confirm the affected version and obtain a minimal synthetic reproduction. Public reports must omit confidential work directories and research figures.
3. Create a focused branch. Add a regression test for substantive bugs, make the smallest appropriate fix and run Windows CI.
4. Changes to Origin/PowerPoint automation, TIFF output, geometry, cancellation or dependencies require a licensed desktop regression. Record sanitized outcomes, including skipped objects and input preservation.
5. Submit a pull request with validation. The owner approves merges. Main branch protection and required checks are repository settings; the project does not use automatic merging.
6. Use semantic versioning: patch for fixes, minor for compatible features, major for breaking input/output or workflow changes. Keep engine/GUI version text, changelog and release notes consistent.
7. Follow [Releasing](RELEASING.md). A reviewed change to the source release configuration triggers source packaging only after checks pass. Windows binaries have additional distribution checks.

Dependabot proposes dependency and GitHub Actions updates. Review advisories and run `pip-audit` before release; updates are not automatically merged. Test new Origin/Office versions when licensed machines are available, and distinguish tested versions from unverified combinations.

Support targets the latest stable release. There is no guaranteed response time or long-term support period. Previous releases remain available; backports require a maintainer decision.

中文：问题通过 Issue 模板登记，修改使用独立分支，测试后提交 PR，由项目所有者确认合并。涉及原生导出和依赖的修改还需在授权 Windows 环境中回归。源码发布与 EXE 发布分开验证，旧版本保留，修复通过新版本提供。
