# CP09A AutoLith Refactoring Plan

Classification: Internal
Baseline: CP09A-REPOSITORY-INVENTORY-20261001-115947

## Phase 1 - Freeze and map
- Freeze the inventory baseline and preserve Summary.json, Files.csv and Files_SHA256.csv.
- Externalize all source, output, threshold and NULL settings.
- Replace hard-coded legacy paths with the new GeoAi root contract.

## Phase 2 - Separate responsibilities
- Isolate source discovery, schema validation, compositing, argmax, comparison, export and validation modules.
- Keep Spence and GeoAI processing independent until the comparison stage.
- Standardize run-scoped output folders and manifests.

## Phase 3 - Migrate safely
- Execute the current baseline and refactored workflow against the same read-only inputs.
- Compare row counts, interval coverage, derived classes, QA/QC and SHA256 outputs.
- Mark every difference as explained, accepted or blocking.

## Phase 4 - Promote
- Promote only after parser, smoke, regression and real-data validation PASS.
- Update launchers, README, manifest, changelog and deployment records.
- Archive the superseded implementation only after rollback validation.

## Decision gates
- Gate 1: path and schema contracts validated.
- Gate 2: numerical and classification equivalence validated.
- Gate 3: governance and rollback evidence complete.
- Gate 4: operational launcher promoted.
