import yaml
import logging
from pathlib import Path
from openpyxl import load_workbook
from src.logger_config import setup_logger
from src.validators import ExcelValidators
from src.merger import ExcelMerger


class Pipeline:
    def __init__(self, config_path="config.yaml"):
        with open(config_path, "r", encoding="utf-8") as file:
            self.cfg = yaml.safe_load(file)

        setup_logger()
        self.data_dir = Path(self.cfg["paths"]["input_dir"])
        self.output_file = self.cfg["paths"]["output_file"]
        self.merger = ExcelMerger(self.output_file)

    def _process_single_file(self, file_path: Path):
        try:
            wb = load_workbook(file_path, data_only=False)
            ws = wb.active
            cfg_v = self.cfg.get("validation", {})

            required_columns = cfg_v.get("required_columns", [])

            if not required_columns:
                logging.error(
                    f"{file_path.name} REJECTED: No required columns defined in config"
                )
                wb.close()
                return

            struct_errors = ExcelValidators.check_columns(ws, required_columns)
            if struct_errors:
                logging.error(f"{file_path.name} REJECTED: {struct_errors[0]}")
                wb.close()
                return

            warnings = []
            warnings.extend(ExcelValidators.check_sequence(ws))
            warnings.extend(ExcelValidators.check_font_style(ws, cfg_v))
            warnings.extend(ExcelValidators.check_x_logic(ws, cfg_v))

            if warnings:
                logging.warning(f"{file_path.name} merged with warnings:")
                for w in warnings:
                    logging.warning(f"  - {w}")
            else:
                logging.info(f"{file_path.name} PASSED and added to merge")
            self.merger.append_data(ws, file_path.name)
            wb.close()

        except Exception as e:
            logging.error(f"{file_path.name} CRASHED: {e}")

    def run(self):
        logging.info("Starting Pipeline with strict validation")

        for file in self.data_dir.rglob("*.xlsx"):
            if file.name.startswith("~") or file.name == self.output_file:
                continue
            if "archive" in [p.lower() for p in file.parts]:
                continue

            self._process_single_file(file)

        self.merger.save()
        logging.info("Pipeline finished")


if __name__ == "__main__":
    Pipeline().run()
