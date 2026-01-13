class Day:
    def __init__(self, pattern):
        self.pattern = pattern

    def __str__(self):
        return str(self.pattern)

    def __repr__(self):
        return self.__str__()