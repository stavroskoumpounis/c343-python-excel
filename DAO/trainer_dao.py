from DTO import Trainer


def add_trainers(sheet, *trainers):
    for trainer in trainers:
        sheet.append((trainer.email, trainer.name, trainer.phone))


def delete_trainers(sheet, *trainer_emails):
    for email in trainer_emails:
        for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row):
            if row[0].value == email:
                sheet.delete_rows(row[0])
                break


def update_trainers(sheet, *trainers):
    for trainer in trainers:
        for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row):
            if row[0].value == trainer.email:
                row[1].value = trainer.name
                row[2].value = trainer.phone
                break


def get_trainer_by_email_id(sheet, email_id):
    for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row):
        if row[0].value == email_id:
            trainer = Trainer(email_id, row[1].value, row[2].value)
            return trainer
