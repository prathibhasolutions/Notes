from src.models import School, Department, Student, Teacher, Administrator, Course
from src.services import SchoolService, ReportService
from src.storage import JsonFileStorage
from src.utils import generate_id


class SchoolApplication:
    """Command-line interface and orchestration layer."""

    def __init__(self, data_dir='data'):
        self.storage = JsonFileStorage(data_dir)
        self.school = self._load_or_create_school()
        self.service = SchoolService(self.school)

    def _load_or_create_school(self):
        school_data = self.storage.load_school()
        if school_data:
            school = School(school_data.get('school_id', 'SCH-001'), school_data.get('name', 'Smart Academy'))
            for dept in school_data.get('departments', []):
                department = Department(dept['department_id'], dept['name'])
                school.add_department(department)
            for student in school_data.get('students', []):
                s = Student(
                    person_id=student.get('person_id', generate_id('PER')),
                    name=student.get('name', 'Unnamed Student'),
                    email=student.get('email', 'student@example.com'),
                    age=student.get('age', 18),
                    student_id=student.get('student_id', generate_id('STU')),
                    course_ids=student.get('course_ids', [])
                )
                school.add_student(s)
            for teacher in school_data.get('teachers', []):
                t = Teacher(
                    person_id=teacher.get('person_id', generate_id('PER')),
                    name=teacher.get('name', 'Unnamed Teacher'),
                    email=teacher.get('email', 'teacher@example.com'),
                    age=teacher.get('age', 36),
                    employee_id=teacher.get('employee_id', generate_id('EMP')),
                    department=teacher.get('department', 'Academic'),
                    teacher_id=teacher.get('teacher_id', generate_id('TCH')),
                    subjects=teacher.get('subjects', [])
                )
                school.add_teacher(t)
            for admin in school_data.get('administrators', []):
                a = Administrator(
                    person_id=admin.get('person_id', generate_id('PER')),
                    name=admin.get('name', 'Admin'),
                    email=admin.get('email', 'admin@example.com'),
                    age=admin.get('age', 30),
                    employee_id=admin.get('employee_id', generate_id('EMP')),
                    department=admin.get('department', 'Administration'),
                    admin_id=admin.get('admin_id', generate_id('ADM'))
                )
                school.add_administrator(a)
            for course in school_data.get('courses', []):
                c = Course(
                    course_id=course.get('course_id', generate_id('CRS')),
                    title=course.get('title', 'Course'),
                    duration_weeks=course.get('duration_weeks', 8),
                    credits=course.get('credits', 2),
                    teacher_id=course.get('teacher_id')
                )
                school.add_course(c)
            return school
        else:
            school = School('SCH-001', 'Smart Academy')
            department = Department('DEP-001', 'Computer Science')
            school.add_department(department)
            return school

    def save_school(self):
        data = {
            'school_id': self.school.school_id,
            'name': self.school.name,
            'departments': [
                {'department_id': d.department_id, 'name': d.name}
                for d in self.school.departments.values()
            ],
            'students': [
                {
                    'person_id': s.person_id,
                    'student_id': s.student_id,
                    'name': s.name,
                    'email': s.email,
                    'age': s.age,
                    'course_ids': s.course_ids,
                }
                for s in self.school.students.values()
            ],
            'teachers': [
                {
                    'person_id': t.person_id,
                    'teacher_id': t.teacher_id,
                    'employee_id': t.employee_id,
                    'name': t.name,
                    'email': t.email,
                    'age': t.age,
                    'department': t.department,
                    'subjects': t.subjects,
                }
                for t in self.school.teachers.values()
            ],
            'administrators': [
                {
                    'person_id': a.person_id,
                    'admin_id': a.admin_id,
                    'employee_id': a.employee_id,
                    'name': a.name,
                    'email': a.email,
                    'age': a.age,
                    'department': a.department,
                }
                for a in self.school.administrators.values()
            ],
            'courses': [
                {
                    'course_id': c.course_id,
                    'title': c.title,
                    'duration_weeks': c.duration_weeks,
                    'credits': c.credits,
                    'teacher_id': c.teacher_id,
                }
                for c in self.school.courses.values()
            ],
            'enrollments': [
                {
                    'enrollment_id': e.enrollment_id,
                    'student_id': e.student_id,
                    'course_id': e.course_id,
                    'status': e.status,
                }
                for e in self.school.enrollments.values()
            ],
        }
        self.storage.save_school(data)

    def _menu(self):
        print('\n====================================')
        print('  SCHOOL MANAGEMENT SYSTEM')
        print('====================================')
        print('1. Register Student')
        print('2. Register Teacher')
        print('3. Register Administrator')
        print('4. Add Course')
        print('5. Enroll Student')
        print('6. Add Mark')
        print('7. View Dashboard')
        print('8. View Reports')
        print('9. Save Data')
        print('0. Exit')
        print('====================================')

    def _register_student(self):
        person_id = generate_id('PER')
        name = input('Student name: ')
        email = input('Student email: ')
        age = int(input('Student age: '))
        student_id = generate_id('STU')
        student = Student(person_id, name, email, age, student_id)
        self.service.register_student(student)
        print('Student registered:', student.name)

    def _register_teacher(self):
        person_id = generate_id('PER')
        name = input('Teacher name: ')
        email = input('Teacher email: ')
        age = int(input('Teacher age: '))
        employee_id = generate_id('EMP')
        department = input('Department: ')
        teacher_id = generate_id('TCH')
        subjects = input('Subjects (comma separated): ').split(',')
        teacher = Teacher(person_id, name, email, age, employee_id, department, teacher_id, [s.strip() for s in subjects])
        self.service.register_teacher(teacher)
        print('Teacher registered:', teacher.name)

    def _register_admin(self):
        person_id = generate_id('PER')
        name = input('Administrator name: ')
        email = input('Administrator email: ')
        age = int(input('Administrator age: '))
        employee_id = generate_id('EMP')
        department = input('Department: ')
        admin_id = generate_id('ADM')
        admin = Administrator(person_id, name, email, age, employee_id, department, admin_id)
        self.service.register_admin(admin)
        print('Administrator registered:', admin.name)

    def _add_course(self):
        course_id = generate_id('CRS')
        title = input('Course title: ')
        duration_weeks = int(input('Duration weeks: '))
        credits = int(input('Credits: '))
        teacher_id = input('Teacher ID (or blank): ')
        course = Course(course_id, title, duration_weeks, credits, teacher_id or None)
        self.service.add_course(course)
        print('Course added:', course.title)

    def _enroll_student(self):
        student_id = input('Student ID: ')
        course_id = input('Course ID: ')
        enrollment = self.service.enroll_student(student_id, course_id)
        if enrollment:
            print('Enrollment created:', enrollment)
        else:
            print('Invalid enrollment details')

    def _add_mark(self):
        student_id = input('Student ID: ')
        course_id = input('Course ID: ')
        mark = float(input('Mark: '))
        student = self.school.get_student(student_id)
        if student:
            student.add_mark(course_id, mark)
            print('Mark added for', student.name)
        else:
            print('Student not found')

    def _dashboard(self):
        dashboard = self.service.get_dashboard()
        print('Dashboard Report')
        print('Total Students:', dashboard['total_students'])
        print('Average Mark:', dashboard['average_mark'])
        print('Student Names:', ', '.join(dashboard['student_names']))
        print('Courses:', ', '.join([course.title for course in dashboard['courses']]))

    def _reports(self):
        report = self.service.get_report_by_grade()
        print('Student Grade Report')
        for name, grade, avg in report:
            print(f'{name}: {grade} | Avg {avg}')

    def run(self):
        while True:
            self._menu()
            choice = input('Choose any option: ')

            if choice == '1':
                self._register_student()
            elif choice == '2':
                self._register_teacher()
            elif choice == '3':
                self._register_admin()
            elif choice == '4':
                self._add_course()
            elif choice == '5':
                self._enroll_student()
            elif choice == '6':
                self._add_mark()
            elif choice == '7':
                self._dashboard()
            elif choice == '8':
                self._reports()
            elif choice == '9':
                self.save_school()
                print('Data saved')
            elif choice == '0':
                self.save_school()
                print('Goodbye')
                break
            else:
                print('Invalid option')
