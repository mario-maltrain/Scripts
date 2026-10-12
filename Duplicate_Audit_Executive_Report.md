# CP09A AutoLith Duplicate Audit Executive Report

Classification: Internal

## Executive recommendation

Do not delete additional data automatically.

The repository cleanup already removed duplicate copy folders supported by complete folder-signature evidence and preserved the valid repository inventory baseline.

The next step is a read-only duplicate audit outside explicitly protected source paths, followed by modular refactoring without changing source data.

## Baseline

- Baseline RUN_ID: CP09A-REPOSITORY-INVENTORY-20261001-115947
- Audit mode: READ_ONLY
- Delete operations: Disabled
- Source modifications: Disabled

## Protected paths

The following paths are excluded from duplicate reporting and must not be deleted:

1. C:/Python/GeoAi/00_Base_Datos/GeoAi/03_Original_Data_Exports/01_Raw_Data/LAS_LEP
2. C:/Python/GeoAi/00_Base_Datos/GeoAi/03_Original_Data_Exports/00_Export/Reblocked/LAS LEP 2026-09-28
3. C:/Python/GeoAi/00_Base_Datos/GeoAi/03_Original_Data_Exports/01_Raw_Data/LEP_Confidence
4. C:/Python/GeoAi/00_Base_Datos/GeoAi/03_Original_Data_Exports/01_Raw_Data/AutoLith_Confidence
5. C:/Python/GeoAi/00_Base_Datos/Spence/00_DataBase

## Audit outputs

- Protected_Paths.csv
- Excluded_Files_Register.csv
- Duplicate_Groups_Outside_Protected.csv
- Duplicate_Files_Outside_Protected.csv
- Duplicate_Summary_By_Folder.csv
- Summary.json
- Summary.txt
- Execution.log
- SHA256SUMS.txt

## Refactoring recommendation

1. Freeze the authoritative source contract.
2. Replace hardcoded source paths with contract-driven path resolution.
3. Add source adapters for Spence, LAS LEP, LEP Confidence and AutoLith Confidence.
4. Separate validation, compositing, argmax, comparison and QA/QC.
5. Treat Master Export as derived and regenerable.
6. Perform side-by-side validation against the frozen baseline.
7. Do not promote the refactored workflow until parser, smoke, regression and real-data validations pass.
