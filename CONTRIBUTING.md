# Contributing

Thanks for helping improve this PRODES pipeline.

## Local Setup

1. Install Python 3.11 or newer.
2. Run `python setup_env.py`.
3. Activate the environment.
4. Run `python 00_pipeline.py --steps 1` for a small smoke test.

## Development Notes

- Keep generated data, presentations, images, and reports out of Git.
- Use `PRODES_HOME` when testing outside the default `workspace/` folder.
- Prefer small, focused commits.
- Run `python -m py_compile *.py` before opening a pull request.

## Data Provider

The pipeline downloads public data from INPE TerraBrasilis. Cite INPE/PRODES
and follow the data provider terms when publishing derived outputs.
