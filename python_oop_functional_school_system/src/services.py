from functools import reduce
from src.models import Student, Teacher, Administrator, Course, Enrollment
from src.utils import calculate_average


class SchoolService:
    """Business logic service layer."""

    def __init__(self, school):
        self.school = school

    def register_student(self, student: Student):
        self.school.add_student(student)
        return student

    def register_teacher(self, teacher: Teacher):
        self.school.add_teacher(teacher)
        return teacher

    def register_admin(self, admin: Administrator):
        self.school.add_administrator(admin)
        return admin

    def add_course(self, course: Course):
        self.school.add_course(course)
        return course

    def enroll_student(self, student_id: str, course_id: str):
        student = self.school.get_student(student_id)
        course = self.school.get_course(course_id)
        if student and course:
            enrollment = Enrollment(f"ENR-{len(self.school.enrollments) + 1}", student_id, course_id)
            self.school.add_enrollment(enrollment)
            student.add_course(course_id)
            return enrollment
        return None

    def get_dashboard(self):
        students = list(self.school.students.values())
        courses = list(self.school.courses.values())
        teachers = list(self.school.teachers.values())

        student_names = list(map(lambda s: s.name, students))
        all_marks = [s.get_average_mark() for s in students]
        average_mark = calculate_average(all_marks)
        total_students = len(students)

        return {
            'student_names': student_names,
            'teachers': teachers,
            'courses': courses,
            'total_students': total_students,
            'average_mark': average_mark,
        }

    def get_report_by_grade(self):
        students = list(self.school.students.values())
        result = []
        for student in students:
            avg = student.get_average_mark()
            if avg >= 90:
                grade = 'A+'
            elif avg >= 75:
                grade = 'A'
            elif avg >= 60:
                grade = 'B'
            elif avg >= 40:
                grade = 'C'
            else:
                grade = 'F'
            result.append((student.name, grade, avg))
        return sorted(result, key=lambda item: item[2], reverse=True)

    def generate_class_list(self):
        return [student.name for student in self.school.students.values()]


class ReportService:
    """Functional report generation service."""

    def __init__(self, students, courses):
        self.students = students
        self.courses = courses

    def student_summary(self):
        summaries = []
        for student in self.students:
            summary = {
                'student_id': student.student_id,
                'name': student.name,
                'average': student.get_average_mark(),
            }
            summaries.append(summary)
        return summaries

    def course_summary(self):
        return list(map(lambda course: {'course_id': course.course_id, 'title': course.title}, self.courses))

    def high_performance_students(self):
        return list(filter(lambda student: student.get_average_mark() >= 75, self.students))
