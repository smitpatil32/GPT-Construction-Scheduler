import openai
import os
import ast
import win32com.client
import datetime

my_api_key = os.getenv("CHAT_GPT_API_KEY")

openai.api_key = my_api_key

# response = openai.ChatCompletion.create(
#   model="gpt-3.5-turbo",
#   messages=[
#         {"role": "system", "content": "You are an expert construciton project scheduler."},
#         {"role": "user", "content": "Give me a list of positive affirmations"},
#     ]
# )

x ="""[['Task name', 'Duration', 'Start date', 'End date'], 
['Site Preparation', 2, 'Jul 1, 2021', 'Jul 2, 2021'], 
['Foundation', 7, 'Jul 5, 2021', 'Jul 13, 2021'], 
['Brickwork', 14, 'Jul 14, 2021', 'Jul 30, 2021'], 
['Door Installation', 2, 'Aug 2, 2021', 'Aug 3, 2021'], 
['Window Installation', 2, 'Aug 4, 2021', 'Aug 5, 2021'], 
['Painting', 5, 'Aug 6, 2021', 'Aug 12, 2021']]"""

tasks_list = ast.literal_eval(x)

Project_App = win32com.client.Dispatch("MSProject.Application")

Project_App.Visible =True

pj =Project_App.Projects.Add()

for task in tasks_list[1:]:
    startdate =datetime.datetime.strptime(task[2],"%b %d, %Y")
    enddate = datetime.datetime.strptime(task[3],"%b %d, %Y")
    duration = task[1]
    task_adder = pj.Tasks.Add(task[0])
    task_adder.Duration = duration
    task_adder.Start = startdate
    task_adder.Finish =enddate



pj.SaveAs(os.path.join(os.getcwd(),"example000.mpp") )
Project_App.Quit()


