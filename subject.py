class Subject:
    _registry = {}
    _id_counter = 0

    def __init__(self, name):
        self.name = name
        self.id = Subject._id_counter
        Subject._registry[self.id] = self
        Subject._id_counter += 1
        self.priority_offline = False
        self.difficulty = 5
        self.max_stack = 1
        self.preferred_stack = 1
        self.max_per_day = 3

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