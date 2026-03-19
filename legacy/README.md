# Legacy Modules

This folder stores modules that were removed from the active SpecCleanse pipeline.

- `style_cleaner.py`
- `deep_cleaner.py`

These components were archived because the current LLM-prep workflow only needs shallow content removal (specifier notes, copyright boilerplate, hidden text, SpecAgent references, and editorial artifacts). The style and deep passes changed XML structure/metadata without meaningful benefit to extracted text.

They are retained for possible future use in a separate DOCX optimization workflow.
