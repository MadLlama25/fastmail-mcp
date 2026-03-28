# Changelog

## 1.9.0 (2026-03-27)


### Features

* add CalDAV fallback for calendar tools ([736ce42](https://github.com/dreamiurg/fastmail-mcp/commit/736ce426f28f4709821d2964e4c57341090ab4ee))
* add edit_draft and send_draft tools ([8e8c61c](https://github.com/dreamiurg/fastmail-mcp/commit/8e8c61c6f4e08462c4952bdc712dfcccd799902e))
* add send parameter to reply_email for draft support ([b4393f7](https://github.com/dreamiurg/fastmail-mcp/commit/b4393f711142f1101b0ea58e10d68c7df82efd03))
* **cli:** add npx-from-GitHub support (bin + prepare + shebang)\n\n- Add bin mapping to dist/index.js and shebang to CLI entry\n- Add prepare script to build TS on GitHub installs (no npm publish)\n- Enforce Node &gt;=18 via engines\n- Document npx usage defaulting to main in README\n- Add CI smoke to validate npx github: repo works on Node 18/20/22\n\nDXT packaging unchanged; manifest and release workflow remain compatible. ([ff5ef3b](https://github.com/dreamiurg/fastmail-mcp/commit/ff5ef3bf9521a10f15b39cb95dfe7d322e89ca3b))


### Bug Fixes

* prevent duplicate drafts from AI tool retries ([66f8362](https://github.com/dreamiurg/fastmail-mcp/commit/66f8362b643da1dadcf5fb31508928f09688ee6a))
* remove dead code, add input guards, and fix double parseInt ([9593b02](https://github.com/dreamiurg/fastmail-mcp/commit/9593b02c8952fe1286085cb5d11a2fdfeec6c7ef))
* resolve content serialization errors after SDK v1.x upgrade ([#27](https://github.com/dreamiurg/fastmail-mcp/issues/27)) ([aac1a92](https://github.com/dreamiurg/fastmail-mcp/commit/aac1a92bda2b8a7a414fe1d70d5219b76c731ee3))
* surface path validation errors in download_attachment ([1cb3ea6](https://github.com/dreamiurg/fastmail-mcp/commit/1cb3ea6c600c01fb5d85d2cbe7f4203ed79ebacf))
* update updateDraft/sendDraft to use getMethodResult() ([dc3f9b4](https://github.com/dreamiurg/fastmail-mcp/commit/dc3f9b45a0944544d8bf178eff8100ae9fe565ac))
* upgrade MCP SDK to v1.x and add null safety to JMAP responses ([#15](https://github.com/dreamiurg/fastmail-mcp/issues/15)) ([#20](https://github.com/dreamiurg/fastmail-mcp/issues/20)) ([7f117db](https://github.com/dreamiurg/fastmail-mcp/commit/7f117db064419bf6403b08d84bf5e1882f4d5786))
