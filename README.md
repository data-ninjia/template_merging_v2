# Excel Data Engineering Pipeline

Python-based tool designed for automated consolidation of technical Excel workbooks with integrated structural validation and formatting checks.

## 🚀 Overview

This project was built to solve the common issue of inconsistent data structures when merging multiple Excel files. It ensures **Data Integrity** by performing a strict structural audit before allowing any data to enter the final merge.

## ✨ Key Features

- **Strict Schema Validation:** Detects missing or extra columns to prevent data shifting.
- **Data Sequence Audit:** Verifies chronological or sequential order (e.g., S/N column).
- **Formatting Compliance:** Identifies non-standard fonts, sizes, and specific business logic markers (e.g., 'X' markers in technical specs).
- **Modular Architecture:** Separated concerns (Validators, Merger, Orchestrator) for easy scalability.
- **Config-Driven:** All validation rules and paths are managed via `YAML`, no hardcoding.
- **Detailed Logging:** Provides clear, actionable logs for rejected files and formatting warnings.

## 🛠️ Tech Stack

- **Language:** Python 3.x
- **Libraries:** `openpyxl` (Excel manipulation), `PyYAML` (Configuration)
- **Environment:** Docker-ready, modular project structure.

## 📁 Project Structure

```text
├── main.py              # Orchestrator: manages the pipeline flow
├── validators.py        # Logic: independent validation methods
├── merger.py            # Engine: handles data & style transfer
├── logger_config.py     # Utilities: professional logging setup
├── config.example.yaml  # Template for configuration
└── requirements.txt     # Project dependencies

