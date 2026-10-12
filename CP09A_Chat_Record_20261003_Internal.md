# CP09A - Registro estructurado del chat 2026-10-03
Clasificación: BHP Internal

## Flujo de la sesión
1. Extracción de AutoLith desde ZIP del export 13-07 (150 LAS) y conversión a CSV: PASS.
2. Motor de fuentes v1.0.0 a v1.0.4: v1.0.0 error de parser; v1.0.2 error .Count; v1.0.3 AutoLith vacío; v1.0.4 READY (12 jobs).
3. Compositación GeoAI 12 jobs: PASS_WITH_WARNINGS. LEP_CONFIDENCE es duplicado de LAS_LEP.
4. Corrección del usuario: la base Spence está en 00_Base_Datos\Spence\00_DataBase\00_TO_COMPOSITE; el job SPENCE_MASTER era el AutoLogger del vendor.
5. Inventario Spence: 14 tablas, Windows-1252 en tablas DET, 30 filas malformadas.
6. Reglas del usuario: compositar solo sondajes del vendor; geología y assays con HOLEID/FROM/TO; texto libre excluido.
7. Lista de sondajes del vendor: 154 (150 con AutoLith).
8. Delta entre exports del vendor: v1.0.0 con falso positivo LEP; v1.0.1 correcto, 853 MAJOR AutoLogger.
9. 10 sondajes SNCT 2025 fuera del COLLAR Spence; presentes en MLP2026 (01_Adhoc_export). Warning previo a compositar implementado.
10. Compositor Spence v1.0.0 falló (NumericColumns vacío); v1.0.1 compositó: CP09A-SPENCE-COMPOSITE-20261003-025200, PASS_WITH_WARNINGS, 11 tablas, 33 jobs, 319,7 s.
11. Actualización de estándares: PowerShell v2.2.0, checklist v2.2.0 (63 MUST), registro Step12 v1.1.0.

## Decisiones del usuario
- Solo se compositan sondajes analizados por el vendor.
- Assays de Spence también se compositan.
- Warning antes de compositar si hay sondajes del vendor fuera del COLLAR.
- Ruta autorizada de collares nuevos: C:\Python\GeoAi\00_Base_Datos\Spence\01_Adhoc_export.
- Todo cambio entre exports del vendor queda en historial visible.

## Requires Validation
- Solapes, intervalo inválido y empates en Spence.
- Combinación de exports 13-07 y 28-09 antes de COMPARE.
- Unidades AU, MO, ICP.
- Exclusión de TipoLitoBrxhMat1/2.
