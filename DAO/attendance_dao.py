def mark_absent(sheet, trainee_ids):
    for row in sheet.iter_rows(min_row=4, max_row=sheet.max_row):
        if row[0].value in trainee_ids:
            row[1].value = 0
