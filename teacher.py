import re

class Teacher:
    name_pattern = r'^[А-ЯІЇЄҐа-яіїєґ]+(?:-[А-ЯІЇЄҐа-яіїєґ]+)* [А-ЯІЇЄҐ]\.[А-ЯІЇЄҐ]\.$'
    _registry = {}
    _id_counter = 0

    def __init__(self, name):
        if not bool(re.fullmatch(self.name_pattern, name)):
            print(f"Warning: Teacher name '{name}' format suspicious.")
        self.name = name
        self.id = Teacher._id_counter
        Teacher._registry[self.id] = self
        Teacher._id_counter += 1
        self.can_offline = True
        self.wants_windows = False
        self.max_online_lessons_from_underground = 2
        self.travel_time = 2

    def __str__(self):
        return self.name

    def __repr__(self):
        return self.__str__()

    @classmethod
    def get_by_id(cls, t_id):
        return cls._registry.get(t_id)

    @classmethod
    def reset_registry(cls):
        cls._registry = {}
        cls._id_counter = 0