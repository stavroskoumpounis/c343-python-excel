import os
from datetime import datetime

import openpyxl

from Employee import *
from Course import Course
from DAO import *

# Your existing code for file creation, reading, and writing here
filename = "python_excel.xlsx"

# Create a new Excel workbook with required sheets if it doesn't exist
if not os.path.exists(filename):
    wb = openpyxl.Workbook()
    # Create and set headers for the ListOfTrainees sheet
    trainees_sheet = wb.active
    trainees_sheet.title = "ListOfTrainees"
    headers = ["id", "Name", "Email", "Course", "Background/degree", "WorkExperience"]
    trainees_sheet.append(headers)

    # Create and set headers for the MappingCourseTrainee sheet
    mapping_course_trainee = wb.create_sheet("MappingCourseTrainee")
    headers = ["CourseId", "TraineeId"]
    mapping_course_trainee.append(headers)

    # Create and set headers for the CourseDetails sheet
    course_details_sheet = wb.create_sheet("CourseDetails")
    headers = ["CourseId", "Description"]
    course_details_sheet.append(headers)

    # Create and set headers for the ListOfTrainers sheet
    trainers_sheet = wb.create_sheet("ListOfTrainers")
    headers = ["Email", "Name", "Phone"]
    trainers_sheet.append(headers)

    # Create and set headers for the MappingCourseTrainee sheet
    mapping_course_trainer = wb.create_sheet("MappingCourseTrainer")
    headers = ["CourseId", "TrainerId"]
    mapping_course_trainer.append(headers)

    manager_sheet = wb.create_sheet("ListOfManagers")
    headers = ["Email", "Name", "Phone", "Based"]
    manager_sheet.append(headers)

    # Create and set headers for the MappingCourseTrainee sheet
    mapping_course_manager = wb.create_sheet("MappingCourseManager")
    headers = ["CourseId", "ManagerId"]
    mapping_course_manager.append(headers)

    wb.save(filename)
else:
    wb = openpyxl.load_workbook(filename)
    trainees_sheet = wb['ListOfTrainees']
    mapping_course_trainee = wb['MappingCourseTrainee']
    mapping_course_manager = wb['MappingCourseManager']
    mapping_course_trainer = wb['MappingCourseTrainer']
    course_details_sheet = wb['CourseDetails']
    trainers_sheet = wb['ListOfTrainers']
    manager_sheet = wb['ListOfManagers']

if not os.path.exists("attendance_book.xlsx"):
    attendance_book = openpyxl.Workbook()
    attendance_book.save("attendance_book.xlsx")
else:
    attendance_book = openpyxl.load_workbook("attendance_book.xlsx")


def create_attendance_sheet(date, start, end, course_id):
    attendance_sheet = attendance_book.create_sheet(date)
    headers = ["Date", "StartTime", "EndTime", "CourseId"]
    attendance_sheet.append(headers)
    attendance_sheet.append((date, start, end, course_id))
    attendance_sheet.append(("TraineeId", "Attended"))
    trainee_ids = course_dao.get_all_trainees_for_course(mapping_course_trainee, course_id)
    if trainee_ids:
        for trainee_id in trainee_ids:
            attendance_sheet.append((trainee_id, 1))
    attendance_book.save("attendance_book.xlsx")


def display_trainees(trainees):
    print("ids of trainees")
    print("-------------------")
    for trainee in trainees:
        print(f"{trainee.id} {trainee.name} {'P' if trainee.attendance else 'A'}")
    print("-------------------")


def record_session_attendance():
    date = input("Enter date (press Enter for current date): ")
    if not date:
        date = datetime.today().strftime('%B%d_%Y')

    start_time = input("Enter start time: ")
    end_time = input("Enter end time: ")
    course_id = input("Enter course ID: ")

    trainer_id = course_dao.get_trainer_for_course(mapping_course_trainer, course_id)
    trainer = trainer_dao.get_trainer_by_email_id(trainers_sheet, trainer_id)
    print(f"Trainer: {trainer.name}")


    # TODO: alter absentees
    trainee_ids = course_dao.get_all_trainees_for_course(mapping_course_trainee, course_id)
    trainees = [trainee_dao.get_trainees_by_id(trainees_sheet, trainee_id)[0] for trainee_id in trainee_ids]
    for trainee in trainees:
        trainee.attendance = True

    display_trainees(trainees)

    absent_ids = input("Enter IDs of absent trainees (comma-separated): ").split(',')
    absent_ids = [int(id.strip()) for id in absent_ids if id.strip()]

    for trainee in trainees:
        if trainee.id in absent_ids:
            trainee.attendance = False

    display_trainees(trainees)

    save = input("Save and submit? (y/n): ")
    if save.lower() == 'y':
        create_attendance_sheet(date, start_time, end_time, course_id)
        attendance_sheet = attendance_book[date]
        attendance_dao.mark_absent(attendance_sheet, absent_ids)


def main_menu():
    while True:
        print("\nMain Menu:")
        print("1. Trainees")
        print("2. Trainers")
        print("3. Managers")
        print("4. Record session attendance")
        print("5. Exit")
        choice = int(input("Enter your choice: "))

        if choice == 1:
            trainee_menu()
        elif choice == 2:
            trainer_menu()
        elif choice == 3:
            manager_menu()
        elif choice == 4:
            record_session_attendance()
        elif choice == 5:
            break
        else:
            print("Invalid choice, try again.")


def trainee_menu():
    while True:
        print("\nTrainee Menu:")
        print("1. Add trainee")
        print("2. Update trainee")
        print("3. Delete trainee")
        print("4. Back to main menu")
        choice = int(input("Enter your choice: "))

        if choice == 1:
            # Add trainee implementation
            pass
        elif choice == 2:
            # Update trainee implementation
            pass
        elif choice == 3:
            # Delete trainee implementation
            pass
        elif choice == 4:
            break
        else:
            print("Invalid choice, try again.")


def trainer_menu():
    # Similar implementation as trainee_menu
    pass


def manager_menu():
    # Similar implementation as trainee_menu
    pass


if __name__ == "__main__":
    main_menu()
