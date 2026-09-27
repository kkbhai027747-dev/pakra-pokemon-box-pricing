# Repository boundaries

This repository owns the local Pokémon box quotation utility only. Start with README.md, collect.py and test_collect.py.

- Preserve pricing rules, Chinese result keys and workbook/JSON output contracts unless the task explicitly changes them.
- Work only on this repository's copy. Do not replace the original computer's running project, change scheduled tasks or start a daily loop as part of ordinary development.
- The utility has no Shopify or WeChat connection. Do not add live store writes or account access to a cleanup task.
- Keep supplier workbooks, config.json, data/, output/, virtual environments and credentials out of Git. Never store or print tokens, cookies or client secrets in source, logs or ordinary .env files.
- Safe regression command, when Python 3.10+ and the existing openpyxl dependency are available: python -B -m unittest -v test_collect. Tests use temporary synthetic workbooks.
- SOURCE-MANIFEST.json records historical provenance; its old paths are not runtime dependencies.
