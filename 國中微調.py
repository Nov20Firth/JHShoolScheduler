# -*- coding: utf-8 -*-
"""驗證瀏覽器拖曳結果，並輸出標示更動位置的 Excel 課表。"""

import json
import sys
from collections import Counter, defaultdict
from pathlib import Path

from Functions import export_to_excel, load_data_file
from Parameters import ALL_SLOTS, DAYS, PERIODS, PHYSICAL_EDUCATION_SUBJECT, ROOM_BY_SUBJECT


if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")


FORMAT_NAME = "junior-high-timetable-adjustment-v3"
RECORD_FIELDS = {
    "id", "teacher", "subject", "klass", "day", "period", "role", "locked_reason",
    "binding_set_id", "binding_group_id", "binding_mode",
}
IDENTITY_FIELDS = (
    "teacher", "subject", "klass", "role", "locked_reason",
    "binding_set_id", "binding_group_id", "binding_mode",
)


def find_adjustment_file(argument=None):
    if argument:
        path = Path(argument).expanduser().resolve()
        return path if path.is_file() else None

    project_dir = Path(__file__).resolve().parent
    search_dirs = [project_dir, Path.home() / "Downloads"]
    candidates = []
    for directory in search_dirs:
        if directory.is_dir():
            for pattern in ("3b_data*.json", "微調資料*.json"):
                candidates.extend(path for path in directory.glob(pattern) if path.is_file())
    return max(candidates, key=lambda path: path.stat().st_mtime) if candidates else None


def read_payload(path):
    try:
        with path.open(encoding="utf-8-sig") as file:
            payload = json.load(file)
    except (OSError, json.JSONDecodeError) as error:
        raise ValueError(f"無法讀取微調資料：{error}") from error
    if payload.get("format") != FORMAT_NAME:
        raise ValueError("微調資料格式不正確，請由本版的暫定課表.html 重新儲存。")
    return payload


def parse_records(value, label):
    if not isinstance(value, list):
        raise ValueError(f"{label}不是課程清單。")
    records = []
    seen_ids = set()
    for index, item in enumerate(value, start=1):
        if not isinstance(item, dict) or not RECORD_FIELDS.issubset(item):
            raise ValueError(f"{label}第 {index} 筆課程欄位不完整。")
        record = {field: item[field] for field in RECORD_FIELDS}
        if record["id"] in seen_ids:
            raise ValueError(f"{label}有重複課程編號：{record['id']}")
        seen_ids.add(record["id"])
        if record["role"] not in {"R", "S"}:
            raise ValueError(f"{label}第 {index} 筆課程類型不正確。")
        if record["day"] not in DAYS or record["period"] not in PERIODS:
            raise ValueError(f"{label}第 {index} 筆課程時段不正確。")
        records.append(record)
    return records


def validate_record_changes(original, adjusted):
    issues = []
    before = {record["id"]: record for record in original}
    after = {record["id"]: record for record in adjusted}
    if set(before) != set(after):
        return ["課程編號集合已被改變；不能新增或刪除課程。"]

    for record_id, old in before.items():
        new = after[record_id]
        if any(old[field] != new[field] for field in IDENTITY_FIELDS):
            issues.append(f"課程 {record_id} 的班級、科目、教師或鎖定狀態被改變。")
            continue
        moved = (old["day"], old["period"]) != (new["day"], new["period"])
        auto_binding_move = (
            old["binding_mode"] == "AUTO"
            and bool(old["binding_set_id"])
            and bool(old["binding_group_id"])
        )
        if moved and old["role"] != "R" and not auto_binding_move:
            issues.append(f"特生課 {old['klass']}班／{old['subject']} 不能由微調介面移動。")
        if moved and old["locked_reason"] and not auto_binding_move:
            issues.append(
                f"{old['klass']}班 {old['subject']}（週{old['day']}第{old['period']}節）"
                f"是{old['locked_reason']}，不能移動。"
            )
    return issues


def validate_schedule(records, courses, teacher_unavailable, fixed_assignments):
    issues = []
    expected_counts = Counter()
    for course in courses:
        expected_counts[course.key] += course.hours
    actual_counts = Counter(
        (record["teacher"], record["subject"], record["klass"], record["role"])
        for record in records
    )
    if actual_counts != expected_counts:
        issues.append("課程堂數或課程內容與 input.txt 不一致。")

    teacher_slots = defaultdict(list)
    class_slots = defaultdict(list)
    room_slots = defaultdict(list)
    course_days = defaultdict(int)
    pe_days = defaultdict(set)
    for record in records:
        teacher = record["teacher"]
        subject = record["subject"]
        klass = record["klass"]
        day = record["day"]
        period = record["period"]
        role = record["role"]
        teacher_slots[teacher, day, period].append(record)
        class_slots[klass, day, period].append(record)
        if (day, period) in teacher_unavailable.get(teacher, set()):
            issues.append(f"{teacher} 在週{day}第{period}節屬於不可排時段。")
        if role == "R":
            course_days[teacher, subject, klass, day] += 1
            room = ROOM_BY_SUBJECT.get(subject)
            if room:
                room_slots[room, day, period].append(record)
            if subject == PHYSICAL_EDUCATION_SUBJECT:
                pe_days[klass].add(DAYS.index(day))

    for (teacher, day, period), rows in teacher_slots.items():
        if len(rows) > 1:
            classes = "、".join(record["klass"] for record in rows)
            issues.append(f"教師衝堂：{teacher} 週{day}第{period}節同時教授 {classes}班。")
    for (klass, day, period), rows in class_slots.items():
        if len(rows) > 1:
            issues.append(f"班級衝堂：{klass}班週{day}第{period}節有 {len(rows)} 堂課。")
    for (room, day, period), rows in room_slots.items():
        if len(rows) > 1:
            issues.append(f"專科教室衝突：{room} 週{day}第{period}節有 {len(rows)} 個班使用。")
    fixed_daily_limits = Counter(
        (teacher, subject, klass, day)
        for teacher, subject, klass, day, period in fixed_assignments
    )
    for (teacher, subject, klass, day), count in course_days.items():
        limit = max(1, fixed_daily_limits[teacher, subject, klass, day])
        if count > limit:
            issues.append(
                f"同科同日重複：{klass}班的{subject}在週{day}排了 {count} 堂，"
                f"允許上限為 {limit} 堂。"
            )
    for klass, day_indices in pe_days.items():
        if any(index + 1 in day_indices for index in day_indices):
            issues.append(f"體育連續兩天：{klass}班的體育排在相鄰上課日。")

    record_keys = {
        (record["teacher"], record["subject"], record["klass"], record["day"], record["period"])
        for record in records if record["role"] == "R"
    }
    for fixed in fixed_assignments:
        if fixed not in record_keys:
            teacher, subject, klass, day, period = fixed
            issues.append(f"固定課被移動：{klass}班 {subject}／{teacher}，原為週{day}第{period}節。")
    return issues


def validate_binding_schedule(payload, original, adjusted, binding_sets):
    issues = []
    schedule = payload.get("binding_schedule")
    if not isinstance(schedule, dict):
        return ["微調資料缺少綁課群組資訊。"]

    expected_groups = {
        group.group_id: (binding, group)
        for binding in binding_sets
        for group in binding.groups
    }
    if set(schedule) != set(expected_groups):
        return ["綁課群組與 input.txt 不一致，請重新產生暫定課表。"]

    original_by_group = defaultdict(list)
    adjusted_by_group = defaultdict(list)
    for record in original:
        if record["binding_group_id"]:
            original_by_group[record["binding_group_id"]].append(record)
    for record in adjusted:
        if record["binding_group_id"]:
            adjusted_by_group[record["binding_group_id"]].append(record)

    original_slots_by_set_day = defaultdict(list)
    adjusted_slots_by_set_day = defaultdict(list)

    for group_id, value in schedule.items():
        if not isinstance(value, dict):
            issues.append(f"綁課群組 {group_id} 的資料格式不正確。")
            continue
        day = value.get("day")
        period = value.get("period")
        assigned = value.get("assigned")
        binding, group = expected_groups[group_id]
        if day not in DAYS or period not in PERIODS or not isinstance(assigned, list):
            issues.append(f"綁課群組 {group_id} 的時段或科目資料不正確。")
            continue
        if value.get("set_id") != binding.set_id or value.get("mode") != group.mode:
            issues.append(f"綁課群組 {group_id} 的組別或安排方式與 input.txt 不一致。")
        if group.mode == "FIXED" and (day, period) != group.fixed_slot:
            issues.append(f"固定綁課群組 {group_id} 的時段與 input.txt 不一致。")
        assigned_pairs = []
        for pair in assigned:
            if not isinstance(pair, list) or len(pair) != 2:
                issues.append(f"綁課群組 {group_id} 的班級科目資料不正確。")
                continue
            assigned_pairs.append(tuple(pair))
        if {klass for klass, subject in assigned_pairs} != set(binding.classes):
            issues.append(f"綁課群組 {group_id} 沒有涵蓋原本所有班級。")
        allowed_subjects = {subject for subject, hours in binding.subjects}
        if any(subject not in allowed_subjects for klass, subject in assigned_pairs):
            issues.append(f"綁課群組 {group_id} 出現不屬於該群組的科目。")

        original_members = original_by_group.get(group_id, [])
        adjusted_members = adjusted_by_group.get(group_id, [])
        if not original_members or len(original_members) != len(adjusted_members):
            issues.append(f"綁課群組 {group_id} 的課程成員不完整。")
            continue
        original_slots = {(record["day"], record["period"]) for record in original_members}
        adjusted_slots = {(record["day"], record["period"]) for record in adjusted_members}
        if len(original_slots) != 1 or len(adjusted_slots) != 1:
            issues.append(f"綁課群組 {group_id} 沒有完整同步移動。")
            continue
        original_slot = next(iter(original_slots))
        adjusted_slot = next(iter(adjusted_slots))
        if adjusted_slot != (day, period):
            issues.append(f"綁課群組 {group_id} 的課程時段與儲存資料不一致。")
        original_pairs = {
            (record["klass"], record["subject"])
            for record in original_members if record["role"] == "R"
        }
        if original_pairs != set(assigned_pairs):
            issues.append(f"綁課群組 {group_id} 的普通班科目與原始課表不一致。")
        if group.mode == "FIXED" and adjusted_slot != original_slot:
            issues.append(f"固定綁課群組 {group_id} 不能移動。")
        if group.mode == "AUTO" and adjusted_slot[0] != original_slot[0]:
            issues.append(f"系統綁課群組 {group_id} 只能在同一天內互換。")
        original_slots_by_set_day[binding.set_id, original_slot[0]].append(original_slot[1])
        adjusted_slots_by_set_day[binding.set_id, adjusted_slot[0]].append(adjusted_slot[1])

    for key, original_periods in original_slots_by_set_day.items():
        if sorted(original_periods) != sorted(adjusted_slots_by_set_day.get(key, [])):
            set_id, day = key
            issues.append(f"綁課組 {set_id} 在週{day}的原有時段被改變。")
    return issues


def to_result_rows(records):
    return [
        (record["teacher"], record["subject"], record["klass"],
         record["day"], record["period"], record["role"])
        for record in records
    ]


def changed_slot_maps(original, adjusted):
    before = {record["id"]: record for record in original}
    class_slots = defaultdict(set)
    teacher_slots = defaultdict(set)
    for new in adjusted:
        old = before[new["id"]]
        if (old["day"], old["period"]) == (new["day"], new["period"]):
            continue
        if new["role"] == "R":
            class_slots[new["klass"]].update({
                (old["day"], old["period"]),
                (new["day"], new["period"]),
            })
        teacher_slots[new["teacher"]].update({
            (old["day"], old["period"]),
            (new["day"], new["period"]),
        })
    return dict(class_slots), dict(teacher_slots)


def deserialize_binding_schedule(value):
    return {
        group_id: (
            item["day"],
            item["period"],
            [tuple(pair) for pair in item["assigned"]],
        )
        for group_id, item in value.items()
    }


def next_output_path(directory):
    number = 1
    while True:
        candidate = directory / f"微調{number}.xlsx"
        if not candidate.exists():
            return candidate
        number += 1


def main():
    argument = sys.argv[1] if len(sys.argv) > 1 else None
    adjustment_path = find_adjustment_file(argument)
    if not adjustment_path:
        print("找不到微調資料.json。請先在暫定課表.html 按下「儲存微調資料」。")
        return 1
    print(f"讀取微調資料：{adjustment_path}")

    try:
        payload = read_payload(adjustment_path)
        original = parse_records(payload.get("original_result"), "原始課表")
        adjusted = parse_records(payload.get("adjusted_result"), "微調課表")
    except ValueError as error:
        print(f"\n{error}")
        return 1

    project_dir = Path(__file__).resolve().parent
    input_path = project_dir / "input.txt"
    if not input_path.is_file():
        print(f"找不到輸入資料：{input_path}")
        return 1
    data = load_data_file(input_path)

    issues = []
    issues.extend(validate_record_changes(original, adjusted))
    issues.extend(validate_schedule(adjusted, data["courses"], data["teacher_unavailable"], data["fixed_assignments"]))
    issues.extend(validate_binding_schedule(payload, original, adjusted, data["binding_sets"]))
    issues = list(dict.fromkeys(issues))
    if issues:
        print("\n微調後發現硬限制衝突，未產生 Excel：")
        for issue in issues:
            print(f"  - {issue}")
        return 1

    moved = sum(
        (old["day"], old["period"]) != (new["day"], new["period"])
        for old, new in zip(
            sorted(original, key=lambda record: record["id"]),
            sorted(adjusted, key=lambda record: record["id"]),
        )
    )
    if moved == 0:
        print("沒有偵測到任何課程更動，未產生 Excel。")
        return 1

    class_slots, teacher_slots = changed_slot_maps(original, adjusted)
    output_path = next_output_path(project_dir)
    binding_schedule = deserialize_binding_schedule(payload["binding_schedule"])
    exported = export_to_excel(
        to_result_rows(adjusted),
        data["school_info"],
        binding_schedule,
        output_path,
        highlighted_class_slots=class_slots,
        highlighted_teacher_slots=teacher_slots,
    )
    if not exported:
        return 1
    print(f"\n微調驗證完成，共移動 {moved} 堂課。")
    print(f"已輸出：{output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
