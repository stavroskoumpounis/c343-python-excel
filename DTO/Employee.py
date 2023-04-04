class Employee:
    def __init__(self, name, email):
        self.name = name
        self.email = email

    def __str__(self):
        return f"Name: {self.name}, Email: {self.email}"


class Trainee(Employee):
    def __init__(self, trainee_id, name, email, course, background, work_exp):
        super().__init__(name, email)
        self.id = trainee_id
        self.course = course
        self.background = background
        self.work_exp = work_exp

    def __str__(self):
        return f"Trainee - ID: {self.id}, {super().__str__()}, Course: {self.course}, Background: {self.background}, Work Experience: {self.work_exp} years"


class Trainer(Employee):
    def __init__(self, email, name, phone):
        super().__init__(name, email)
        self.phone = phone

    def __str__(self):
        return f"Trainer - {super().__str__()}, Phone: {self.phone}"


class Manager(Employee):
    def __init__(self, email, name, phone, based, courses):
        super().__init__(name, email)
        self.phone = phone
        self.based = based
        self.courses = courses if courses else []

    def __str__(self):
        courses = ', '.join([course.description for course in self.courses])
        return f"Manager - {super().__str__()}, Based: {self.based}, Courses: [{courses}]"
