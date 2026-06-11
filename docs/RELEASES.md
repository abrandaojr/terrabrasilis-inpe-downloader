# Releases

This repository uses Git tags for release points.

## Creating a Release

1. Update `CHANGELOG.md`.
2. Run `python -m py_compile run_pipeline.py setup_env.py scripts/*.py prodes_pipeline/*.py`.
3. Commit the changes.
4. Create an annotated tag, for example:

   ```bash
   git tag -a v1.1.0 -m "v1.1.0 portable pipeline release"
   git push origin main --tags
   ```

5. Create a GitHub Release from the tag and paste the matching changelog entry.

## Package Status

`pyproject.toml` provides local package metadata for helper modules. The numbered
pipeline scripts are kept as runnable scripts rather than console entry points.
GitHub Packages publication is intentionally not automated until package or
container distribution is needed.
