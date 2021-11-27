import pandas as pd
from pprint import pprint
import json

df = pd.read_excel("google_form_reply.xlsx")
# print(df)

columns = df.columns
num_of_students = df.shape[0] - 1

# Name
names = df[columns[2]]
ids = df[columns[3]]

# Student ID
students = {}
for idx in range(num_of_students):
    students[str(ids[idx+1])] = {
        "Name": names[idx+1],
        "Scores": [],
        "Wrong Question": [],
        "Corrections": [],
        "Total": 0
    }

# Questions
score_4_question = [5,5,5,4,5,4,4,2,2,2,2,2,2,2,2,2,4,4,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2]
for question_idx, column in enumerate(columns[4:]):
    correct_answer = df[column][0]
    correct_answer = str(correct_answer).replace("\u2028", "")
    student_answers = df[column][1:]
    
    for student_idx, answer in enumerate(student_answers, 1):
        student_id = str(ids[student_idx])
        answer = str(answer).replace('”', '"')
        answer = str(answer).replace('“', '"')
        answer = str(answer).replace('” ', '"')
        answer = str(answer).replace('“ ', '"')
        answer = str(answer).replace('" ', '"')
        answer = str(answer).replace(' "', '"')
        answer = answer.strip()
        
        if str(answer) in str(correct_answer) or str(correct_answer) in str(answer):
            students[student_id]["Scores"].append(score_4_question[question_idx])
        else:
            students[student_id]["Scores"].append(0)
            students[student_id]["Wrong Question"].append(column)
            students[student_id]["Corrections"].append("{} >> {}".format(str(answer), str(correct_answer)))

# Sum up
for student_id in students:
    students[student_id]["Total"] = sum(students[student_id]["Scores"])

# Average
print("Average:", sum(students[student_id]["Scores"]) / num_of_students * 100)

with open('results.json', 'w', encoding='utf-8') as f:
    json.dump(students, f, ensure_ascii=False, indent=4)
