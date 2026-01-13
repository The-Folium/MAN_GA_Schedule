class DataBase:
    def __init__(self):
        self.teacher_dict = {}
        self.subject_dict = {}
        self.lesson_list = []
        self.school_class_dict = {}
        self.subtable_coords = []

    def get_class_load(self, class_id):
        counter = 0
        for lesson in self.lesson_list:
            if lesson.school_class.name == class_id:
                counter += (0.5 if lesson.is_blinking else 1)
        return counter