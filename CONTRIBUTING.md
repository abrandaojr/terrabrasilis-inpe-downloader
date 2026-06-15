# Contributing

Thanks for helping improve this PRODES pipeline.

## Local Setup

1. Install Python 3.11 or newer.
2. Run `python setup_env.py`.
3. Activate the environment.
4. Run `python run_pipeline.py --steps 1` for a small smoke test.

## Development Notes

- Keep generated data, presentations, images, and reports out of Git.
- Use `PRODES_HOME` when testing outside the default `workspace/` folder.
- Prefer small, focused commits.
- Keep PRODES analyses anchored to the 2008 analytical base year; earlier years
  may be shown only as illustrative context.
- Treat `VS_`, `vegetacao_secundaria`, and `floresta_secundaria` layers as
  secondary vegetation, not PRODES deforestation.
- Run `python -m compileall -q run_pipeline.py setup_env.py scripts prodes_pipeline tests` before opening a pull request.
- Run `python -m unittest discover -s tests` for regression tests.

## Data Provider

The pipeline downloads public data from INPE TerraBrasilis. Cite INPE/PRODES
and follow the data provider terms when publishing derived outputs.
