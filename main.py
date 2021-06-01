import json
import pprint
import re
import win32api
import win32con
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select

CONST_IS_EXAM = "ללא בחינה"


def click(x, y):
    win32api.SetCursorPos((x, y))
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, x, y, 0, 0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, x, y, 0, 0)


def scrap_courses(department_number, department_name):
    try:
        this_department_course_list = []
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "BTN_RESET"))
        ).click()
        courses_current_department = {}
        department = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.ID, "on_course_department"))
        )
        department.send_keys(department_number)
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "GOPAGE2"))
        ).click()
        courses_table = driver.find_element_by_id('courseTable')
        rows = courses_table.find_elements(By.TAG_NAME, "tr")  # get all of the rows in the table
        rows = rows[1:]
        for row in rows:
            # Get the columns (all the column 2)
            col = row.find_elements(By.TAG_NAME, "td")
            course_name = col[2].text
            course_number = col[0].text
            open_in = col[1].text

            this_department_course_list.append({
                "department_number": department_number,
                "department_name": department_name,
                "course_name": course_name,
                "course_number": course_number,
                "open_in": open_in
            })

        driver.find_element_by_class_name('styled-button-2').click()
        department = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.ID, "on_course_department"))
        )
        department.clear()
        return this_department_course_list

    except Exception as e:
        print(f"Exception in scrap courses function : {str(e)}")
        department = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.ID, "on_course_department"))
        )
        department.clear()
        pass


if __name__ == '__main__':
    courses = []
    courses_extended = {}
    departments = []
    course_page_template = 'https://bgu4u.bgu.ac.il/pls/scwp/!app.gate?app=ann&lang=he&step=3&st=s&popup=cns' \
                           '&rn_course_department={department}&rn_course_degree_level={' \
                           'course_degree_level}&rn_course={' \
                           'course_number}&rn_course_details=1&rn_course_ins=0&rn_year={year}&rn_semester={' \
                           'semester_number} '

    driver = webdriver.Chrome('C:\\Users\\Nir\\PycharmProjects\\seleniumCourseScraper\\chromedriver.exe')
    driver.get('https://bgu4u.bgu.ac.il/pls/scwp/!app.gate?app=ann')
    frame = driver.find_element_by_name('main')
    driver.switch_to.frame(frame)
    driver.implicitly_wait(5)
    click(500, 450)
    departments_list_selector = Select(driver.find_element_by_id('on_course_department_list'))
    options = departments_list_selector.options
    for index in range(0, len(options) - 1):
        department_number_name_text = str(options[index].text).split(' - ')
        if len(department_number_name_text) == 2:
            departments.append((department_number_name_text[0], department_number_name_text[1]))

    try:
        for department_number_name in departments:
            try:
                department_number = department_number_name[0]
                department_name = department_number_name[1]
                if department_number.startswith('0'):
                    department_number = department_number.lstrip('0')
                result = scrap_courses(department_number, department_name)
                if result is not None:
                    for item in result:
                        courses.append(item)
                time.sleep(2)
            except Exception as e:
                print(f"Exception in main : {str(e)}")
                department = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.ID, "on_course_department"))
                )
                department.clear()
                continue

        # for department in list(courses.keys()):
        #     try:
        #         print(department)
        #         courses_extended[department] = {}
        #         for course in list(courses[department].keys()):
        #             print(course)
        #             _, course_degree_level, course_number = courses[department][course]['course_number'].split('.')
        #             year, semester_number = courses[department][course]['open_in'].split('-')
        #             url = course_page_template.format(department=department,
        #                                               course_degree_level=course_degree_level,
        #                                               course_number=course_number,
        #                                               year=year, semester_number=semester_number)
        #             driver.get(url)
        #             frame = driver.find_element_by_name('main')
        #             driver.switch_to.frame(frame)
        #             course_unordered_list = driver.find_elements_by_tag_name('li')
        #
        #             credit_points = str(course_unordered_list[4].text).split(':')[1].rstrip('.00')
        #             test = str(course_unordered_list[12].text)
        #             li_list = driver.find_elements_by_class_name('BlackInput')
        #
        #             time = ''
        #             time_re = re.compile('(([01][0-9]|2[0-3]):([0-5][0-9]))')
        #             match_time = time_re.findall(str(li_list[3].text))
        #             if match_time is not None and len(match_time) == 2:
        #                 time = match_time[0][0] + '-' + match_time[1][0]
        #
        #             day = ''
        #             day_time = ''
        #             day_re = re.compile('(יום [א|ב|ג|ד|ה|ו])')
        #             match_day = day_re.findall(str(li_list[3].text))
        #             if len(match_day) > 0:
        #                 day_time = match_day[0] + ' ' + time
        #
        #             print(day_time)
        #             courses_extended[department][course] = {
        #                 "course_number": courses[department][course]['course_number'],
        #                 # "open_in": courses[department][course]['open_in'],
        #                 "credit_points": credit_points,
        #                 "test_exist": not (CONST_IS_EXAM in test),
        #                 "professor name": li_list[2].text
        #                 # "time": day_time
        #             }
        #
        #             print('\n')
        #             print('_' * 69)
        #
        #     except Exception as e:
        #         print(f"Exception in scraping specific course in department : {department}\n{str(e)}")
        #         department = WebDriverWait(driver, 10).until(
        #             EC.visibility_of_element_located((By.ID, "on_course_department"))
        #         )
        #         department.clear()
        #         continue


    finally:
        courses_to_save = json.dumps(courses, ensure_ascii=False, indent=4, sort_keys=True)
        extened_courses_to_save = json.dumps(courses_extended, ensure_ascii=False, indent=4, sort_keys=True)
        with open('courses.json', 'wb') as f:
            f.write(courses_to_save.encode('utf-8'))
        with open('extended_courses.json', 'wb') as f:
            f.write(extened_courses_to_save.encode('utf-8'))
        driver.implicitly_wait(5)
        driver.quit()
