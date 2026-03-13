class ExcelValidators:
    @staticmethod
    def check_columns(ws, required_columns: list) -> list:
        """Check presence of mandatory columns"""
        errors = []
        actual_columns = [str(cell.value).strip() for cell in ws[1] if cell.value]

        actual_set = set(actual_columns)
        required_set = set(required_columns)

        missing = required_set - actual_set
        if missing:
            errors.append(f"MISSING columns: {', '.join(missing)}")

        extra = actual_set - required_set
        if extra:
            errors.append(f"EXTRA columns: {', '.join(extra)}")
        return errors

    @staticmethod
    def check_sequence(ws) -> list:
        errors = []
        sn_col_idx = next((cell.column for cell in ws[1] if cell.value == "S/N"), None)
        if not sn_col_idx:
            return ["Column 'S/N' not found"]

        seq, rows = [], []
        for row in ws.iter_rows(min_row=2, max_col=sn_col_idx, min_col=sn_col_idx):
            if row[0].value is not None:
                try:
                    seq.append(int(row[0].value))
                    rows.append(row[0].row)
                except (ValueError, TypeError):
                    pass

        expected = list(range(1, len(seq) + 1))
        if seq != expected:
            for i, (act, exp) in enumerate(zip(seq, expected)):
                if act != exp:
                    errors.append(
                        f"Sequence error at row {rows[i]}: expected {exp}, got {act}"
                    )
                    break
        return errors

    @staticmethod
    def check_font_style(ws, config: dict) -> list:
        errors = []
        expected_font = config.get("font_name")
        expected_size = config.get("font_size")

        for row in ws.iter_rows(min_row=2):
            for cell in row:
                if cell.value is not None:
                    f = cell.font
                    if (f.name != expected_font) or (f.size != expected_size):
                        return [f"FONT ERROR: Expected {expected_font} {expected_size}"]
        return []

    @staticmethod
    def check_x_logic(ws, config: dict) -> list:
        errors = []
        target_name = config.get("target_col_name")
        target_idx = next(
            (
                c.column
                for c in ws[1]
                if c.value and str(c.value).strip() == target_name
            ),
            None,
        )
        if not target_idx:
            return [f"Logic check skipped: Column '{target_name}' not found"]
        prev_colored = False
        for row in ws.iter_rows(min_row=2):
            fill = row[0].fill
            curr_colored = bool(fill and fill.patternType == "solid")
            if prev_colored and not curr_colored:
                val = ws.cell(row=row[0].row, column=target_idx).value
                if not val or str(val).strip().lower() != "x":
                    errors.append(f"LOGIC ERROR: Missing 'X' at row {row[0].row}")
            prev_colored = curr_colored
        return errors
