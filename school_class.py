class SchoolClass:
    _registry = {}
    _id_counter = 0

    def __init__(self, name):
        self.name = name
        self.id = SchoolClass._id_counter
        SchoolClass._registry[self.id] = self
        SchoolClass._id_counter += 1
        self.priority_offline = None

    def __str__(self):
        return self.name

    def __repr__(self):
        return self.__str__()

    def __hash__(self):
        return hash(self.name)

    @classmethod
    def get_by_id(cls, t_id):
        return cls._registry.get(t_id)

    @classmethod
    def reset_registry(cls):
        cls._registry = {}
        cls._id_counter = 0