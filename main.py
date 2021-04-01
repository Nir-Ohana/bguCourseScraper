import json
import pprint

import win32api
import win32con
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select


def click(x, y):
    win32api.SetCursorPos((x, y))
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, x, y, 0, 0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, x, y, 0, 0)


def scrap_courses(department_number):
    try:
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

            courses_current_department[course_name] = {
                "course_number": course_number,
                "open_in": open_in
            }

        driver.find_element_by_class_name('styled-button-2').click()
        department = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.ID, "on_course_department"))
        )
        department.clear()
        return courses_current_department

    except Exception as e:
        print(f"Exception in scrap courses function : {str(e)}")
        department = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.ID, "on_course_department"))
        )
        department.clear()
        pass


if __name__ == '__main__':
    courses = {}
    departments = []

    driver = webdriver.Chrome('C:\\Users\\Nir\\PycharmProjects\\seleniumCourseScraper\\chromedriver.exe')
    driver.get('https://bgu4u.bgu.ac.il/pls/scwp/!app.gate?app=ann')
    # department = WebDriverWait(driver, 10).until(
    #     EC.visibility_of_element_located((By.CLASS_NAME, "menuLinkU"))
    # )[0].click()
    # driver.find_element_by_class_name('menuLinkU')[0].click()
    frame = driver.find_element_by_name('main')
    driver.switch_to.frame(frame)
    driver.implicitly_wait(5)
    click(500, 450)
    departments_list_selector = Select(driver.find_element_by_id('on_course_department_list'))
    options = departments_list_selector.options
    for index in range(0, len(options) - 1):
        department_number_name_text = str(options[index].text).split(' - ')
        departments.append(department_number_name_text[0])

    departments = departments[1:]
    try:
        for department_number in departments:
            try:
                if department_number.startswith('0'):
                    department_number = department_number.lstrip('0')
                print(department_number)
                result = scrap_courses(department_number)
                if result is not None:
                    courses.update(
                        {
                            department_number: result
                        }
                    )
                # print(pprint.pformat(courses))
                time.sleep(2)
            except Exception as e:
                print(f"Exception in main : {str(e)}")
                department = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.ID, "on_course_department"))
                )
                department.clear()
                continue

    finally:
        # with open('courses.json', 'w') as f:
        #     json.dump(courses, f, ensure_ascii=False)
        to_save = json.dumps(courses, ensure_ascii=False, indent=4, sort_keys=True)
        with open('courses.json', 'wb') as f:
            f.write(to_save.encode('utf-8'))
        driver.implicitly_wait(5)
        driver.quit()
