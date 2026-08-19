# -*- coding: utf-8 -*-
"""國中排課使用的學制、科目、權重與求解參數。"""

DAYS = ["一", "二", "三", "四", "五"]
PERIODS = [1, 2, 3, 4, 5, 6, 7]
ALL_SLOTS = [(day, period) for day in DAYS for period in PERIODS]
REQUIRED_CLASS_HOURS = len(ALL_SLOTS)

PERIOD_TIMES = {
    1: ("08:25", "09:10"),
    2: ("09:20", "10:05"),
    3: ("10:15", "11:00"),
    4: ("11:10", "11:55"),
    5: ("13:15", "14:00"),
    6: ("14:10", "14:55"),
    7: ("15:15", "16:00"),
}

DEFAULT_SCHOOL_INFO = {
    "school_name": "○○○○○○國民中學",
    "school_year": "○○○",
    "semester": "○",
}
SCHOOL_NAME_SUFFIX = "國民中學"

ROOM_BY_SUBJECT = {
    "生活科技": "生活科技教室",
    "資訊科技": "資訊教室",
    "音樂": "音樂教室",
    "視覺藝術": "美術教室",
}
SUBJECT_ALIASES = {
    "國文": "國語文",
    "英文": "英語文",
    "資科": "資訊科技",
    "資訊": "資訊科技",
    "生科": "生活科技",
    "表藝": "表演藝術",
    "表演": "表演藝術",
    "視藝": "視覺藝術",
    "美術": "視覺藝術",
    "閩南語": "閩南語文",
    "客語": "客語文",
    "原住民族語": "原住民族語文",
    "閩東語": "閩東語文",
    "手語": "臺灣手語",
    "台灣手語": "臺灣手語",
    "新住民語": "新住民語文",
    "自然": "自然科學",
    "健康": "健康教育",
    "健教": "健康教育",
    "地科": "地球科學",
}


def normalize_subject(value):
    subject = str(value).strip()
    return SUBJECT_ALIASES.get(subject, subject)


EXAM_SUBJECTS = {
    "國語文",
    "英語文",
    "數學",
    "自然科學",
    "生物",
    "理化",
    "地球科學",
    "社會",
    "地理",
    "歷史",
    "公民",
}
PHYSICAL_EDUCATION_SUBJECT = "體育"
HOMEROOM_TEACHER_SUFFIX = "導師"

MIN_EXAM_HOURS_FOR_GAP_RULE = 3
MAX_EXAM_PERIODS_PER_DAY = 5

WEIGHT_CROSS_SUBJECT = 100
WEIGHT_EXAM_DAILY_OVERLOAD = 60
WEIGHT_TEACHER_LUNCH = 40
WEIGHT_EXAM_PERIOD_7 = 30
WEIGHT_TEACHER_CONSECUTIVE_THREE = 20
WEIGHT_HOMEROOM_EARLY = 10
WEIGHT_EXAM_TWO_DAY_GAP = 5
WEIGHT_EXAM_PERIOD_5 = 4
WEIGHT_TEACHER_LATE = 3

SOLVER_MAX_TIME_SECONDS = 60
SOLVER_NUM_SEARCH_WORKERS = 8
