def export_to_excel(data, output_path, cell_bboxes):
    logging.debug(f"Exporting data to Excel: {output_path}")
    workbook = Workbook()
    sheet = workbook.active

    # Group bounding boxes by Y coordinates
    rows = {}
    for bbox, cell_value in zip(cell_bboxes, data):
        x1, y1, x2, y2 = bbox
        if y1 not in rows:
            logging.debug(f"Creating new row for Y coordinate: {y1}")
            rows[y1] = []
        rows[y1].append((x1, cell_value))

    # Sort rows by Y coordinates and columns by X coordinates
    sorted_rows = sorted(rows.items())
    for row_num, (y1, cells) in enumerate(sorted_rows, start=1):
        sorted_cells = sorted(cells)
        row_values = [cell_value for x, cell_value in sorted_cells]
        print(f"Row {row_num}: {row_values}")  # Print each row to the terminal
        for col_num, (x1, cell_value) in enumerate(sorted_cells, start=1):
            logging.debug(f"Appending cell value: {cell_value} at row {row_num}, column {col_num}")
            sheet.cell(row=row_num, column=col_num, value=cell_value)

    workbook.save(output_path)
    logging.info(f"Exported data to Excel file: {output_path}")
