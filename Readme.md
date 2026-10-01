# Data Vault Studio, Complete Visualization Edition

## Run on Windows
```powershell
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

The full application includes template download, metadata/relationship validation, Hub/Link/Satellite generation, advanced Satellite splitting, Effectivity Satellites, bulk/table/artifact/column review, quality gates, naming comparison, SQL generation, versioning, exports, and three visualization tabs:

- High-level ER: reviewed Hubs and Links
- Low-level ER: reviewed Hubs, Links, Satellites, columns and optional technical/source details
- Data flow: sources, stage, reviewed Raw Vault and optional Business Vault/Mart/Consumption nodes from `flow_nodes` and `flow_edges`

Diagram previews use Streamlit Graphviz. The export package includes editable `.dot` sources for all three diagrams. If local rendering reports a Graphviz error, install Graphviz for Windows and add its `bin` directory to PATH.
