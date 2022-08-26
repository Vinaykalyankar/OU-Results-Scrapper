from selenium import webdriver
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
import time
from openpyxl import Workbook

driver = webdriver.Chrome(ChromeDriverManager().install())
start=1000001
end=1000010
student_final_result=""
student_name=""
# driver.maximize_window()
driver.get("https://www.osmania.ac.in/res07/20220535.jsp")
wb=Workbook()
sh1=wb.active

students_dict={}
students_list=[]
for i in range(start,end):
    try:
        student_details=[]

        student_sub_grade={}

        driver.implicitly_wait(10)
        driver.find_element(By.XPATH,"//input[contains(@name,'htno')]").send_keys(i)
        driver.find_element(By.XPATH,"//input[contains(@value,'Go')]").click()
        driver.implicitly_wait(10)
        S_details=driver.find_elements(By.XPATH,"//td[contains(@width,'33')]")
        for s1 in S_details:
            print(s1.text)
            student_details.append(s1.text)
        student_name=student_details[1]
        # print("Type of S Detials : ",type(S_details))
        S_Subjects=driver.find_elements(By.XPATH,"//td[contains(@width,'50')]")
        S1_Subjects=[]
        for j in S_Subjects[1:]:
            S1_Subjects.append(j.text)
        S2_Subjects=[]
        for k in S1_Subjects:
            if k=="Result":
                break
            else:
                S2_Subjects.append(k)
        # print(S2_Subjects)

        S_Grades=driver.find_elements(By.XPATH,"//td[contains(@width,'20')]")
        S1_Grades=[]
        for j in S_Grades[1:]:
            S1_Grades.append(j.text)
        # print(S1_Grades)
        student_sub_grade = dict(zip(S2_Subjects, S1_Grades))
        print(student_sub_grade)
        Subject_names=[' INFORMATION SECURITY',' OBJECT ORIENTED SYSTEM DEVELOPMENT',' BIG DATA ANALYTICS',' CLOUD COMPUTING',' SOFTWARE QUALITY AND TESTING',' PROGRAMMING LAB-IX (OOSD)',' PROGRAMMING LAB-X (BIG DATA ANALYTICS)',' PROJECT SEMINAR','DATA MINING','WEB PROGRAMMING','DISTRIBUTED DATABASES','WEB PROGRAMMING (LAB)']
        d12={}
        for key, value in student_sub_grade.items():
            print("key",key)
            if key in Subject_names:
                d12[key] = value
        print("D1: ",d12)

        # S_Results=driver.find_elements(By.XPATH,"//td[contains(@width,'50')]")
        # print(S_Results.text)
        # print(type(S_Results))

        student_final_result=S_Subjects[-1].text
        print(student_final_result)
        OnlyGrades=[]
        for k,v in d12.items():
            OnlyGrades.append(v)
        student_list=[]
        student_list.append(i)
        student_list.append(student_name)

        for i in OnlyGrades:
            student_list.append(i)
        student_list.append(student_final_result)
        print("Stud List: ",student_list)
        # students_dict[i]=[student_name,d12,student_final_result]
        students_list.append(student_list)
        print("Students list: ",students_list)

        time.sleep(2)
    except Exception as e:
        print(e)
        continue;
    print(students_dict)

for x in students_list:
    sh1.append(x)



wb.save("finalRecords121.xlsx")
driver.quit()
