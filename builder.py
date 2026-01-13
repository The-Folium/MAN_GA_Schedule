import parameters as P
from config import lessons_available_per_day
from prepare import SLOT_EMPTY, SLOT_ONLINE, SLOT_OFFLINE, SLOT_TRAVEL, SLOT_UNWANTED


class TimetableBuilder:
    def __init__(self, individual, db, slots_structure):
        self.individual = individual
        self.db = db
        self.slots_structure = slots_structure
        self.teachers_by_id = {t.id: t for t in db.teacher_dict.values()}
        self.subjects_by_id = {s.id: s for s in db.subject_dict.values()}
        self.schedule = {}
        for class_obj in self.db.school_class_dict.values():
            self.schedule[class_obj.id] = [[None] * lessons_available_per_day for _ in range(7)]

        self.teacher_timeline = {}
        for teacher_obj in self.db.teacher_dict.values():
            self.teacher_timeline[teacher_obj.id] = [[None] * lessons_available_per_day for _ in range(7)]
        self.teacher_online_from_school_counters = {
            t_id: 0 for t_id in self.teachers_by_id.keys()
        }
        self.f1_student_hardships = 0.0
        self.f2_didactic_quality = 0.0
        self.f3_teacher_comfort = 0.0
        self.unplaced_lessons_count = 0

    def build(self):
        for lesson_id in self.individual:
            if not self._try_place_lesson(lesson_id):
                self.unplaced_lessons_count += 1
                self.f1_student_hardships += P.P_UNPLACED_LESSON
        self._calculate_post_build_metrics()
        return (self.f1_student_hardships, self.f2_didactic_quality, self.f3_teacher_comfort)

    def _try_place_lesson(self, lesson_id):
        lesson = self.db.lesson_list[lesson_id]
        c_id = lesson.school_class_id
        for compromise_mode in (False, True):
            for slot_idx in range(lessons_available_per_day):
                for day_idx in range(7):
                    slot_type = self.slots_structure[c_id][day_idx][slot_idx]
                    if slot_type == SLOT_UNWANTED and not compromise_mode:
                        continue
                    if self._can_place(lesson, day_idx, slot_idx):
                        self._commit_lesson(lesson, day_idx, slot_idx)
                        return True
        return False

    def _can_place(self, lesson, day, slot):
        c_id = lesson.school_class_id
        t_id = lesson.teacher_id
        if self.schedule[c_id][day][slot] is not None:
            return False
        slot_type = self.slots_structure[c_id][day][slot]
        if slot_type == SLOT_EMPTY or slot_type == SLOT_TRAVEL:
            return False
        teacher = self.teachers_by_id[t_id]
        if slot_type == SLOT_OFFLINE and not teacher.can_offline:
            return False
        if self.teacher_timeline[t_id][day][slot] is not None:
            return False
        target_loc = 'SCHOOL' if slot_type == SLOT_OFFLINE else 'HOME'
        if not self._check_teacher_travel_constraint(teacher, day, slot, target_loc):
            return False
        return True

    def _check_teacher_travel_constraint(self, teacher, day, slot, target_loc):
        timeline = self.teacher_timeline[teacher.id][day]
        travel_needed = teacher.travel_time
        prev_slot = slot - 1
        while prev_slot >= 0:
            status = timeline[prev_slot]
            if status is not None:
                prev_loc = status['loc']
                if prev_loc != target_loc:
                    is_online_from_school_case = (prev_loc == 'SCHOOL' and target_loc == 'HOME')
                    if not is_online_from_school_case:
                        gap = slot - prev_slot - 1
                        if gap < travel_needed:
                            return False
                break
            prev_slot -= 1
        next_slot = slot + 1
        while next_slot < lessons_available_per_day:
            status = timeline[next_slot]
            if status is not None:
                next_loc = status['loc']
                if next_loc != target_loc:
                    gap = next_slot - slot - 1
                    if gap < travel_needed:
                        return False
                break
            next_slot += 1
        return True

    def _commit_lesson(self, lesson, day, slot):
        slot_type = self.slots_structure[lesson.school_class_id][day][slot]
        subject = self.subjects_by_id[lesson.subject_id]
        teacher = self.teachers_by_id[lesson.teacher_id]
        actual_loc = 'HOME'
        if slot_type == SLOT_OFFLINE:
            actual_loc = 'SCHOOL'
        else:
            prev_slot = slot - 1
            while prev_slot >= 0:
                st = self.teacher_timeline[teacher.id][day][prev_slot]
                if st:
                    if st['loc'] == 'SCHOOL':
                        actual_loc = 'SCHOOL'
                    break
                prev_slot -= 1
        self.schedule[lesson.school_class_id][day][slot] = lesson.id
        self.teacher_timeline[teacher.id][day][slot] = {
            'loc': actual_loc,
            'lesson_id': lesson.id
        }

        if slot_type == SLOT_UNWANTED:
            self.f1_student_hardships += P.P_UNWANTED_SLOT

        if subject.priority_offline and slot_type != SLOT_OFFLINE:
            self.f2_didactic_quality += P.P_SUBJECT_MISMATCH

        if actual_loc == 'SCHOOL' and slot_type != SLOT_OFFLINE:
            self.teacher_online_from_school_counters[teacher.id] += 1
            if self.teacher_online_from_school_counters[teacher.id] > teacher.max_online_lessons_from_underground:
                self.f3_teacher_comfort += P.P_EXCESS_ONLINE_FROM_SCHOOL

    def _calculate_post_build_metrics(self):
        for c_id, days in self.schedule.items():
            daily_difficulties = []

            for day_lessons in days:
                gaps = self._count_gaps(day_lessons)
                self.f1_student_hardships += gaps * P.P_STUDENT_GAP
                current_stack_subj = None
                current_stack_count = 0
                day_difficulty = 0
                day_subjects_counter = {}
                for idx, l_id in enumerate(day_lessons):
                    if l_id is None:
                        self.f2_didactic_quality += self._check_stack(current_stack_subj, current_stack_count)
                        current_stack_subj = None
                        current_stack_count = 0
                        continue

                    lesson = self.db.lesson_list[l_id]
                    subject = self.subjects_by_id[lesson.subject_id]
                    day_subjects_counter[subject.id] = day_subjects_counter.get(subject.id, 0) + 1
                    day_difficulty += subject.difficulty
                    if subject.difficulty >= 7:
                        if idx == 0 or idx >= (lessons_available_per_day - 1):
                            self.f2_didactic_quality += P.P_DIFFICULTY_DISTRIBUTION
                    if current_stack_subj == subject.id:
                        current_stack_count += 1
                    else:
                        self.f2_didactic_quality += self._check_stack(current_stack_subj, current_stack_count)
                        current_stack_subj = subject.id
                        current_stack_count = 1
                self.f2_didactic_quality += self._check_stack(current_stack_subj, current_stack_count)
                daily_difficulties.append(day_difficulty)
                for s_id, count in day_subjects_counter.items():
                    subj = self.subjects_by_id[s_id]
                    if count > subj.max_per_day:
                        excess = count - subj.max_per_day
                        self.f2_didactic_quality += excess * P.P_MAX_LESSONS_PER_DAY
            avg_mid = sum(daily_difficulties[1:4]) / 3 if sum(daily_difficulties[1:4]) > 0 else 1
            if daily_difficulties[0] > avg_mid * 1.15:
                self.f2_didactic_quality += P.P_DIFFICULTY_DISTRIBUTION * 2
            if daily_difficulties[4] > avg_mid * 1.15:
                self.f2_didactic_quality += P.P_DIFFICULTY_DISTRIBUTION * 2
        for t_id, days in self.teacher_timeline.items():
            teacher = self.teachers_by_id[t_id]
            for day_slots in days:
                gaps = self._count_gaps(day_slots)
                if gaps > 0:
                    if not teacher.wants_windows:
                        self.f3_teacher_comfort += gaps * P.P_TEACHER_GAP

    def _count_gaps(self, lst):
        has_started = False
        gap_count = 0
        in_gap = False
        last_idx = -1
        for i in range(len(lst) - 1, -1, -1):
            if lst[i] is not None:
                last_idx = i
                break

        if last_idx == -1:
            return 0
        for i in range(last_idx + 1):
            val = lst[i]
            if val is not None:
                has_started = True
                if in_gap:
                    gap_count += 1
                    in_gap = False
            else:
                if has_started:
                    in_gap = True
        return gap_count

    def _check_stack(self, subj_id, count):
        if subj_id is None: return 0
        subject = self.subjects_by_id[subj_id]
        p = 0
        if count > subject.max_stack:
            p += (count - subject.max_stack) * P.P_STACK_VIOLATION
        elif count != subject.preferred_stack:
            p += P.P_STACK_NON_PREFERRED
        return p