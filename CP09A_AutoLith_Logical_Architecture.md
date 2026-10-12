# CP09A AutoLith Logical Architecture

Classification: Internal
Baseline: CP09A-REPOSITORY-INVENTORY-20261001-115947

## 1. Data layer
- 00_Base_Datos/Spence/00_DataBase: authoritative Spence source data, read-only.
- 00_Base_Datos/GeoAi/03_Original_Data_Exports/01_Raw_Data: authoritative GeoAI raw inputs, read-only.
- 00_Base_Datos/CP09A_Output: immutable historical CP09A outputs.

## 2. Application layer
- 01_Aplicacion/CP09A_Release_v2.1/00_Script: launchers, inventory, validation and cleanup tools.
- 01_Source: source adapters and input contracts.
- 02_Model: deterministic AutoLith compositing and argmax logic.
- 03_Output: current application outputs only.
- 04_Manifest: manifests, schemas and SHA256 records.
- 05_Logs: execution logs.
- 06_Backup: controlled rollback assets.
- 07_Archive: superseded releases retained only when required.

## 3. Processing contract
1. Discover and validate source paths.
2. Load Spence and GeoAI sources in read-only mode.
3. Validate HOLE_ID, intervals, probability columns, NULL rules and coverage threshold.
4. Composite each source independently.
5. Derive AutoLith class by argmax on aggregated raw probabilities without normalization.
6. Compare sources only in the comparison stage.
7. Export run-scoped outputs, QA/QC, manifest and SHA256.
8. Promote only validated PASS runs.

## 4. Governance
- Every execution requires a unique RUN_ID.
- Inputs remain immutable.
- Outputs are isolated by run.
- No deletion is allowed without hash or supersession evidence.
- Classification is Internal.
