import os
import re
import copy
import xlrd
import argparse
from pathlib import Path


def list_to_matrix(courses):
    """
    Convert a list of course info to a matrix of course name
    :param courses: a list of course info, the index interpretation of courses are:
           ["课程名称(班级)", "开课院系", "任课教师", "上课时间地点", "校区", "学时", "学分", "容纳人数", "预选人数"]
    :return: a matrix of course name
    """

    # initialize the course matrix
    course_matrix = []
    for i in range(11):  # 一天最多11节课
        row = []
        for j in range(7):  # 一个星期7天
            row.append(None)
        course_matrix.append(row)

    for course_info in courses:
        course_name = course_info[0]

        classtime_and_classroom = course_info[3]
        pattern = re.compile(r"[0-9]{1,2}-[0-9]{1,2}周.*?\)")
        matched_str = pattern.findall(classtime_and_classroom)

        for classtime_and_classroom in matched_str:
            pattern = re.compile(r"[0-9]{1,2}-[0-9]{1,2}周，")
            weeks = pattern.findall(classtime_and_classroom)[0]
            pattern = re.compile(r"\(.*\)")
            classroom = pattern.findall(classtime_and_classroom)[0]

            pattern = re.compile(r"星期..[1-9]-[1-9]{1,2}节")
            matched_str = pattern.findall(classtime_and_classroom)

            for classtime in matched_str:
                day = classtime[:3]

                col_idx = {
                    '星期一': 0,
                    '星期二': 1,
                    '星期三': 2,
                    '星期四': 3,
                    '星期五': 4,
                    '星期六': 5,
                    '星期日': 6
                }[day]

                pattern = re.compile(r"[1-9]-[1-9]{1,2}")
                begin, finish = pattern.findall(classtime)[0].split('-')
                begin = int(begin) - 1
                finish = int(finish)

                for row_idx in range(begin, finish):
                    if not course_matrix[row_idx][col_idx]:
                        course_matrix[row_idx][col_idx] = course_name + "<br />" + weeks + classroom
                    else:
                        course_matrix[row_idx][col_idx] = course_matrix[row_idx][col_idx] + "<br />" + weeks + classroom

    return course_matrix


def write_to_markdown(course_matrix):
    """
    write course matrix to a markdown file
    :param course_matrix: a course matrix
    :return:
    """

    with open("Schedule.md", 'a', encoding="utf-8") as f:
        print("|   | 星期一 | 星期二 | 星期三 | 星期四 | 星期五 | 星期六 | 星期天 |", file=f)
        print("| ---- | ------ | ------ | ------ | ------ | ------ | ------ | ------ |", file=f)

        for row_idx, course_info in enumerate(course_matrix):
            print("| " + str(row_idx + 1), end=' | ', file=f)
            for col_idx in range(7):
                if course_info[col_idx]:
                    print(course_info[col_idx], end=' | ', file=f)
                else:
                    print(' ', end=' | ', file=f)
            print(file=f)

        print(file=f)
        print(file=f)


def scheduling(key, courses_dict, selected_times, selected_courses):
    for course_info in courses_dict[key]:
        classtime_and_classroom = course_info[3]  # 上课时间和地点
        pattern = re.compile(r"星期..[1-9]-[1-9]{1,2}节")
        matched_str = pattern.findall(classtime_and_classroom)

        available = True
        for classtime in matched_str:
            available = available and classtime not in selected_times

        if available:
            selected_times_copy = copy.deepcopy(selected_times)
            selected_courses_copy = copy.deepcopy(selected_courses)

            for classtime in matched_str:
                selected_times_copy.append(classtime)
            selected_courses_copy.append(course_info)

            courses_dict_copy = copy.deepcopy(courses_dict)
            courses_dict_copy.pop(key)
            if len(courses_dict_copy) == 0:
                course_matrix = list_to_matrix(selected_courses_copy)
                write_to_markdown(course_matrix)
            else:
                next_key = list(courses_dict_copy.keys())[0]
                scheduling(next_key, courses_dict_copy, selected_times_copy, selected_courses_copy)


def run():
    parser = argparse.ArgumentParser()
    parser.add_argument("--campus", type=str, default='', help="清水河或沙河")
    parser.add_argument("--time_placeholder", type=str, action="append", default=[], help="时间占位符，设置格式示例：星期三第5-6节")
    parser.add_argument("--excel", type=str, default='', help="课程信息Excel文件所在路径")
    args = parser.parse_args()

    print(args)

    if not os.path.exists(args.excel):
        raise FileNotFoundError

    excel_path = Path(args.excel)
    workbook_obj = xlrd.open_workbook(f"{excel_path}")
    sheet_data = workbook_obj.sheet_by_name("Sheet1")

    courses_dict = {}

    # 表头占了一行，所以是(1, 0)
    previous_course_id = sheet_data.cell(1, 0).value
    try:
        previous_course_id = str(int(previous_course_id))
    except ValueError:
        previous_course_id = str(previous_course_id)

    # row_data索引含义：
    # ["课程编号", "课程名称(班级)", "开课院系", "任课教师", "上课时间地点", "校区", "学时", "学分", "容纳人数", "预选人数"]
    # 表头占了一行，所以从1开始
    for row_idx in range(1, sheet_data.nrows):
        row_data = sheet_data.row_values(row_idx)

        course_id = row_data[0]
        course_time = row_data[4]
        if course_id != '' and course_time != '':

            if args.campus:
                if row_data[5] != args.campus:
                    continue

            try:
                course_id = str(int(course_id))
            except ValueError:
                course_id = str(course_id)
            previous_course_id = course_id
            course_info = row_data[1:6] + [int(row_data[6]), int(row_data[7]), int(row_data[8]), int(row_data[9])]
            if course_id in courses_dict:
                courses_dict[course_id].append(course_info)
            else:
                courses_dict[course_id] = [course_info]
        else:
            if row_data[4] != '':
                courses_dict[previous_course_id][-1][3] += ' ' + row_data[4]

    selected_courses = []
    for k, v in courses_dict.items():
        scheduling(k, copy.deepcopy(courses_dict), copy.deepcopy(args.time_placeholder),
                   copy.deepcopy(selected_courses))


if __name__ == "__main__":
    run()
