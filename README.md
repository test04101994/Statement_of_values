# SOV parser

Statement of Values (SOV) Excel ingestion with AWS Bedrock and configurable field schema.

## Layout

| Path | Purpose |
|------|---------|
| `sov/` | Python package (`engine`, parsing pipeline) |
| `config/` | `config.yaml`, optional `learned_synonyms.yaml`, `embeddings_cache.json` |
| `samples/` | Example `.xlsx` workbooks |
| `scripts/` | `create_sample.py`, `create_sample_addr.py` (generators) |
| `tests/` | Unit tests |
| `main.py` | CLI entry point |
| `output/` | Run artifacts (JSON sidecars; gitignored) |

## Quick start

```bash
python -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
python main.py samples/sample_sov.xlsx
```

Override config: `python main.py file.xlsx --config /path/to/config.yaml`

Run tests: `python -m unittest discover -s tests -v`
