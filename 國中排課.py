# -*- coding: utf-8 -*-
"""國中課表排程器。

輸入資料使用 S、C、F、BK、BT、U：
    S|學校名稱|學年度|學期
    C|老師|科目|堂數|班級
    F|老師|科目|班級|星期|節次
    BK|綁課組|班級1,班級2|科目1:堂數,科目2:堂數
    BT|綁課組|群組|AUTO/FIXED|星期或-|節次或-|特教老師或-|特教科目或-|特生班或-
    U|老師|星期|節次
"""

import sys
from pathlib import Path

from Functions import (
    build_and_solve,
    export_to_adjustment_html,
    load_data_file,
    print_data_summary,
)


def main():
    input_path = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).with_name("input.txt")
    output_path = Path(sys.argv[2]) if len(sys.argv) > 2 else input_path.with_name("暫定課表.html")

    print(f"讀取資料：{input_path}")
    data = load_data_file(input_path)
    print_data_summary(data["courses"], data["binding_sets"])
    print("開始排課，請稍候...")
    solved = build_and_solve(
        data["courses"],
        data["teacher_unavailable"],
        data["fixed_assignments"],
        data["binding_sets"],
    )
    if not solved:
        return 1
    result, binding_schedule = solved
    exported = export_to_adjustment_html(
        result,
        data["school_info"],
        binding_schedule,
        data["binding_sets"],
        data["fixed_assignments"],
        data["teacher_unavailable"],
        output_path,
    )
    if not exported:
        return 1
    print(f"暫定課表已匯出：{output_path}")
    print("請在瀏覽器拖曳課程，完成後按下「儲存微調資料」。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
