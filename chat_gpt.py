import openai
import os
import ast
import win32com.client
import datetime
import win32timezone
from time import sleep
from getpass import getpass

if os.getenv('CHAT_GPT_API_KEY') is None:
    API_KEY = getpass("Input your API key: ")
    openai.api_key = API_KEY

    with open(".env", "a") as file:
        file.write(f"CHAT_GPT_API_KEY={API_KEY}\n")
else:
    openai.api_key = os.getenv('CHAT_GPT_API_KEY')


def get_details_from_user():
    main_string = """With the following 'construction details information' create a schedule for the construction project described. The schedule should include a start date and an end date and breakdown the construction project into specific tasks and mention the time allocated for each task in days. Your response should only represent the schedule in a in a list of lists format. each list inside the main list should contain following columns in this specific order - [Task name, Duration, Start date, end date, predecessor]. The first row should be the title row. and the following rows should contain the data. The "Predecessors" property expects a string where tasks are identified by their ID number and separated by a comma. For example, if task 3 and 5 are predecessors to task 6, it would look like this: "3,5". I do not want any citations or explanations. Use the current date and time for the schedule that you create."""
    details_str = ""
    prjname = input("Project Name: ")
    details_str += f"Project Name: {prjname}; "
    location = input("Project Location: ")
    details_str += f"Project Location: {location}; "
    projtype = input(
        "Project Type (e.g., residential, commercial, industrial): ")
    details_str += f"Project Type: {projtype}; "
    dimensions = input("Dimensions (length x width x height, in meters): ")
    details_str += f"Dimensions (length x width x height, in meters): {dimensions}; "
    material = input(
        "Type of Construction Material (e.g., brick, concrete, wood): ")
    details_str += f"Type of Construction Material (e.g., brick, concrete, wood): {material}; "
    wallthickness = input("Wall Thickness (in meters): ")
    details_str += f"Wall Thickness (in meters): {wallthickness}; "
    numdoors = input("Number of Doors: ")
    details_str += f"Number of Doors: {numdoors}; "
    doordim = input(
        "Door Dimensions (width x height x thickness, in meters): ")
    details_str += f"Door Dimensions (width x height x thickness, in meters): {doordim}; "
    doortype = input("Type of Door (e.g., wooden, metal, glass): ")
    details_str += f"Type of Door (e.g., wooden, metal, glass): {doortype}; "
    numwindows = input("Number of Windows (if any): ")
    details_str += f"Number of Windows: {numwindows}; "
    windowdim = input("Window Dimensions (width x height, in meters): ")
    details_str += f"Window Dimensions (width x height, in meters): {windowdim}; "
    windowtype = input(
        "Type of Window (e.g., sliding, casement, double-hung): ")
    details_str += f"Type of Window (e.g., sliding, casement, double-hung): {windowtype}; "
    painttype = input("Paint Type and Color: ")
    details_str += f"Paint Type and Color: {painttype}; "
    paintdim = input("Paint Thickness (in millimeters): ")
    details_str += f"Paint Thickness (in millimeters): {paintdim}; "
    ceiling = input("Ceiling and Roof Required (Yes/No): ")
    details_str += f"Ceiling and Roof Required: {ceiling}; "
    elecwork = input("Electrical Work Required (Yes/No): ")
    details_str += f"Electrical Work Required: {elecwork}; "
    plumbing = input("Plumbing Work Required (Yes/No): ")
    details_str += f"Plumbing Work Required: {plumbing}; "
    deadline = input("Project Deadline (in weeks): ")
    details_str += f"Project Deadline (in weeks): {deadline}; "

    return main_string+details_str


def get_response_from_ChatGPT(prompt):
    print("\nasking ChatGPT...")
    response = openai.ChatCompletion.create(
        model="gpt-3.5-turbo",

        messages=[
            {"role": "system", "content": "You are an expert construciton project scheduler."},
            {"role": "user", "content": prompt},
        ]
    )
    print("\nGot response from ChatGPT!:")
    print(response.choices[0].message.content)
    return response.choices[0].message.content


def create_mpp_file(response_string):
    # Create a list from the given chatgpt output
    print("\nCreating .MPP file")
    tasks_list = eval(response_string)

    project_name = input("\nEnter the Project Name: ")
    project_name += "_"+datetime.datetime.now().strftime('%m-%d_%H-%M-%p')

    # Initialize the MS Project App
    Project_App = win32com.client.Dispatch("MSProject.Application")
    Project_App.Visible = True
    pj = Project_App.Projects.Add()

    for i, task in enumerate(tasks_list[1:]):
        print(f"creating task {i}")
        startdate = datetime.datetime.strptime(task[2], "%b %d, %Y")
        enddate = datetime.datetime.strptime(task[3], "%b %d, %Y")
        duration = task[1]
        task_adder = pj.Tasks.Add(task[0])
        task_adder.Duration = duration
        task_adder.Start = startdate
        task_adder.Finish = enddate
        if task[4]:  # Assuming 4th index contains the predecessors
        # Assign predecessors.
            task_adder.Predecessors = task[4]

    print(f"\nsaving project as {project_name}.mpp")
    pj.SaveAs(os.path.join(os.getcwd(), f"{project_name}.mpp"))
    # print(f'\nsaved as {os.path.dirname(os.path.join(os.getcwd(),"output", f"{project_name}.mpp"))}')
    Project_App.Quit()


test_output = """[['Task name', 'Duration', 'Start date', 'End date'],
['Site Preparation', 2, 'Jul 1, 2021', 'Jul 2, 2021'],
['Foundation', 7, 'Jul 5, 2021', 'Jul 13, 2021'],
['Brickwork', 14, 'Jul 14, 2021', 'Jul 30, 2021'],
['Door Installation', 2, 'Aug 2, 2021', 'Aug 3, 2021'],
['Window Installation', 2, 'Aug 4, 2021', 'Aug 5, 2021'],
['Painting', 5, 'Aug 6, 2021', 'Aug 12, 2021']]"""

if __name__ == '__main__':
    # user_prompt = get_details_from_user()
    user_prompt = """With the following 'construction details information' create a schedule for the construction project described. The schedule should include a start date and an end date and breakdown the construction project into specific tasks and mention the time allocated for each task in days. Your response should only represent the schedule in a in a list of lists format. each list inside the main list should contain following columns in this specific order - [Task name, Duration, Start date, end date, predecessor]. The first row should be the title row. and the following rows should contain the data. The "Predecessors" property expects a string where tasks are identified by their ID number and separated by a comma. For example, if task 3 and 5 are predecessors to task 6, it would look like this: "3,5". I do not want any citations or explanations. Use the current date and time for the schedule that you create.
 Project Name Sample
# Location Indiana 
# Type: Residential 
# Start Date (e.g., Jul 26, 2023): Jul 26, 2023
# Type of Construction Material (e.g., brick, concrete, wood) Brick and Mortar
# Wall Thickness (in meters): 0.20
# Number of Doors: 1
# Door Dimensions (width x height x thickness, in meters): 2 x 1 x 0.20
# Type of Door (e.g., wooden, metal, glass): Wooden
# Number of Windows (if any): 2
# Window Dimensions (width x height, in meters) 1 x 1
# Type of Window (e.g., sliding, casement, double-hung): Fixed
# Paint Type and Color: Flat and White 
# Paint Thickness (in millimeters): 2
# Ceiling and Roof Required (Yes/No): Yes
# Electrical Work Required (Yes/No): No
# Plumbing Work Required (Yes/No): No
# Project Deadline (in weeks): None"""
    gpt_reponse = get_response_from_ChatGPT(user_prompt)
    # # print(gpt_reponse)
#     gpt_reponse = """[['Task name', 'Duration', 'Start date', 'End date', 'Predecessors'],
# ['Site preparation', 3, 'Jul 26, 2023', 'Jul 29, 2023', ''],
# ['Foundation construction', 5, 'Jul 30, 2023', 'Aug 3, 2023', '1'],
# ['Wall construction', 10, 'Aug 4, 2023', 'Aug 15, 2023', '2'],
# ['Door installation', 2, 'Aug 16, 2023', 'Aug 17, 2023', '3'],
# ['Window installation', 2, 'Aug 16, 2023', 'Aug 17, 2023', '3'],
# ['Painting', 5, 'Aug 18, 2023', 'Aug 22, 2023', '4,5'],
# ['Ceiling and roof installation', 7, 'Aug 23, 2023', 'Aug 29, 2023', '6'],
# ['Final inspection and cleanup', 2, 'Aug 30, 2023', 'Aug 31, 2023', '7']]"""
    create_mpp_file(gpt_reponse)
