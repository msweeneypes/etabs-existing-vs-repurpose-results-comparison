## Project
Viktor.ai app — ETABS model comparison tool
Python, Viktor framework
https://docs.viktor.ai/llms-full.txt

## Run
viktor-cli start

## Structure
- app.py — main Viktor app entry point
- [add other key files as you build]

## Conventions
- Follow Viktor component patterns (views, params, controllers)
- Keep comparison logic in separate modules from UI code

## Data
- ETABS Excel exports, .xlsx format
- Key sheets: Element Forces - Columns/Beams/Braces
- Key columns: OutputCase, CaseType, Frame, Station, P, V2, V3, T, M2, M3
- Combo names follow pattern: "Stl - D2.3-Xa.X) 1.2D + ..."