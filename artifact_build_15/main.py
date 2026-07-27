from __future__ import annotations

import shutil
from pathlib import Path

import build_collective_prep
import build_course_ac
import build_course_full
import build_curriculum
import build_plan
import build_student
import build_teacher_four
import build_teacher_tier

BASE = Path(__file__).resolve().parent
OUTPUT_BASE = BASE / "output"
PACKAGE = OUTPUT_BASE / "开江中学实验学校党委书记15份参考文件优化完善稿"


def main() -> None:
    if OUTPUT_BASE.exists():
        shutil.rmtree(OUTPUT_BASE)
    PACKAGE.mkdir(parents=True, exist_ok=True)

    build_student.build_all(PACKAGE)
    build_teacher_tier.build(PACKAGE)
    build_teacher_four.build(PACKAGE)
    build_curriculum.build(PACKAGE)
    build_collective_prep.build_all(PACKAGE)
    build_course_ac.build_all(PACKAGE)
    build_course_full.build(PACKAGE)
    build_plan.build(PACKAGE)

    docs = sorted(PACKAGE.rglob("*.docx"))
    if len(docs) != 15:
        raise RuntimeError(f"Expected 15 Word documents, found {len(docs)}: {[str(x) for x in docs]}")

    lines = [
        "开江中学实验学校党委书记15份参考文件优化完善稿",
        "",
        "说明：本文件包以学校领导、党委书记发来的原始稿件为基础，保留主体架构和主要内容，进行针对性校正、实操化补充、格式统一和有限度风险调整。其他5份现实问题整改与重点学生工作背景材料未纳入本次修改。",
        "",
        "文件目录：",
    ]
    for index, path in enumerate(docs, 1):
        lines.append(f"{index:02d}. {path.relative_to(PACKAGE)}")
    lines.extend([
        "",
        "建议使用顺序：先由对应分管领导和牵头部门核定处室名称、年度时间、真实数据及本地政策要求；涉及学生处分、教师绩效职称、心理健康、安全、数据和收费事项的，须完成专项审核后发布。",
    ])
    (PACKAGE / "00_文件目录与使用说明.txt").write_text("\n".join(lines), encoding="utf-8")
    print(f"Generated {len(docs)} Word documents in {PACKAGE}")


if __name__ == "__main__":
    main()
