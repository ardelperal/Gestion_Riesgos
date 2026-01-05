# Repository Guidelines

## Project Structure & Module Organization
Exported VBA modules live at the repo root: `.cls` files map to Access classes (e.g., `Proyecto.cls`, `Riesgo.cls`) and `.bas` files hold shared modules such as `Constructor.bas`, `Funciones Generales.bas`, and automation helpers like `ExcelInforme.bas`. Keep UI-centric form code in the `Form_*` classes to isolate logic from presentation, and leave `References.txt` untouched because it documents the required COM libraries. Tests and smoke-check helpers reside in `testing.bas`.

## Build, Test, and Development Commands
Open the Access front-end that consumes these modules (usually `GestionProyectos.accdb`) with Access 2016+. Import modified files via `External Data > New Data Source > From File > Access`, then compile with `Debug > Compile GestionProyectos`. To launch the client with a specific macro from a shell, use `"C:\Program Files\Microsoft Office\root\Office16\MSACCESS.EXE" GestionProyectos.accdb /x AutoExec`. Run data refresh routines from the Immediate window, e.g., `Call Instalacion.ActualizarEntorno` before distributing a release.

## Coding Style & Naming Conventions
Follow `Option Explicit` in every module and indent procedures with four spaces. Use PascalCase for class modules (`RiesgoExterno`, `ProyectoSuministrador`), camelCase for locals (`m_Riesgo`, `sErr`), and keep enum/const prefixes consistent with `EnumSiNo` and `Constantes.bas`. Favor descriptive Spanish names that mirror business terms, and keep public procedures grouped by feature with a short header comment explaining purpose and parameters.

## Testing Guidelines
Add lightweight assertions to `testing.bas`; each function should start with `test_` and return a diagnostic string. Execute tests from the Immediate window (`? test_Riesgo_Dias_Por_Aceptar`) or wire them to a macro so QA can run `DoCmd.RunMacro "RunTests"`. When touching risk workflows, cover both acceptance and withdrawal paths using the provided constructors, and document any required TempVars or seed data in the header comment.

## Commit & Pull Request Guidelines
This snapshot lacks Git metadata, so default to Conventional Commits (`feat:`, `fix:`, `chore:`) with short bilingual subjects when helpful. Reference the Access feature or risk code you touched in the body (e.g., `Ref: R012`). Pull requests should describe schema changes, include screenshots for updated forms, and link to the incident or ticket. List manual test steps (import, compile, smoke test macro) before requesting review.

## Security & Configuration Tips
Respect the COM references listed in `References.txt`; mismatch versions can break JSON parsing (`JsonConverter.bas`) or Excel exports. Never commit credentials—store DSN/user secrets in Access-linked tables outside version control. Before packaging, verify TempVars defaults and clean cached data by running `Constructor.LimpiarCache` (when available) to avoid leaking production entities.