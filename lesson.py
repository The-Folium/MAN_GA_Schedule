from teacher import Teacher
from subject import Subject
from school_class import SchoolClass


class Lesson:
    _id_counter = 0
    def __init__(self, school_class_id, subject_id, teacher_id, is_blinking=False):
        self.id = Lesson._id_counter
        Lesson._id_counter += 1
        self.school_class_id = school_class_id
        self.subject_id = subject_id
        self.teacher_id = teacher_id
        self.is_blinking = is_blinking

    def __str__(self):
        return f"Lesson({SchoolClass.get_by_id(self.school_class_id)}|{Subject.get_by_id(self.subject_id)}|{Teacher.get_by_id(self.teacher_id)})"

    def __repr__(self):
        return self.__str__()

    @classmethod
    def reset_registry(cls):
        cls._id_counter = 0