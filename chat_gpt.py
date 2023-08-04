import openai
import os
import ast
import win32com.client
import datetime
import win32timezone
from time import sleep
from getpass import getpass
import json

# Try to load API key from json file
try:
    with open('api_key.json', 'r') as file:
        data = json.load(file)
        openai.api_key = data.get('CHAT_GPT_API_KEY')
except (FileNotFoundError, KeyError):
    pass

# If API key is not found, ask for input
if openai.api_key is None:
    API_KEY = getpass("Input your API key: ")

    # Provide asterisks as feedback after user input
    print("API key received: " + "*" * len(API_KEY))

    # Save API key into json file
    with open('api_key.json', 'w') as file:
        json.dump({'CHAT_GPT_API_KEY': API_KEY}, file)

    openai.api_key = API_KEY



def get_details_from_user():
    main_string = """With the following 'construction details information' create a schedule for the construction project described. The schedule should include a start date and an end date and breakdown the construction project into specific tasks and mention the time allocated for each task in days. Your response should only represent the schedule in a in a list of lists format. each list inside the main list should contain following columns in this specific order - [Task name, Duration, Start date, end date, predecessor]. The first row should be the title row. and the following rows should contain the data. The "Predecessors" property expects a string where tasks are identified by their ID number and separated by a comma. For example, if task 3 and 5 are predecessors to task 6, it would look like this: "3,5". In addition, you cannot make a task a predecessor of itself. I do not want any citations or explanations. Use the current date and time for the schedule that you create."""
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
    start_date = input("Start Date (e.g., Jul 26, 2023): ")
    details_str +=f"Start Date (e.g., Jul 26, 2023): {start_date}; "
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
            {
                "role": "system",
                "content": "You are an expert Construction Project Manager who specializes in breaking down a project into its consequent project activities and schedule them to create a realistic schedule for the construction project to be completed in optimal time."
            },
            {
                "role": "user",
                "content": prompt
            }
        ],
        temperature=0.7,
        max_tokens=1024,
        top_p=0.8,
        frequency_penalty=0,
        presence_penalty=0
    )

    print("\nGot response from ChatGPT!:")
    print(response.choices[0].message.content)
    return response.choices[0].message.content


def create_mpp_file(response_string):
    # Create a list from the given chatgpt output
    print("\nCreating .MPP file")
    tasks_list = eval(response_string)

    project_name = input("\nEnter the Project Name: ")
    # project_name += "_"+datetime.datetime.now().strftime('%m-%d_%H-%M-%p')

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


test_output = """[['Task name', 'Duration', 'Start date', 'End date', 'Predecessor'],
 ['Site Preparation', 2, 'Jul 26, 2023', 'Jul 27, 2023', ''],
 ['Foundation Construction', 5, 'Jul 28, 2023', 'Aug 2, 2023', '1'],
 ['Wall Construction', 5, 'Aug 3, 2023', 'Aug 8, 2023', '2'],
 ['Door Installation', 2, 'Aug 9, 2023', 'Aug 10, 2023', '3'],
 ['Window Installation', 2, 'Aug 11, 2023', 'Aug 12, 2023', '4'],
 ['Painting', 3, 'Aug 13, 2023', 'Aug 15, 2023', '5'],
 ['Ceiling and Roof Construction', 4, 'Aug 16, 2023', 'Aug 19, 2023', '6'],
 ['Project Completion', 0, 'Aug 19, 2023', 'Aug 19, 2023', '7']]"""

if __name__ == '__main__':
    user_prompt = get_details_from_user()
    #     user_prompt = """With the following 'construction details information' create a schedule for the construction project described. The schedule should include a start date and an end date and breakdown the construction project into specific tasks and mention the time allocated for each task in days. Your response should only represent the schedule in a in a list of lists format. each list inside the main list should contain following columns in this specific order - [Task name, Duration, Start date, end date, predecessor]. The first row should be the title row. and the following rows should contain the data. The "Predecessors" property expects a string where tasks are identified by their ID number and separated by a comma. For example, if task 3 and 5 are predecessors to task 6, it would look like this: "3,5". I do not want any citations or explanations. Use the current date and time for the schedule that you create.
    #  Project Name Sample
    # # Location Indiana
    # # Type: Residential
    # Dimensions (length x width x height, in meters): 4 x 4
    # # Start Date (e.g., Jul 26, 2023): Jul 26, 2023
    # # Type of Construction Material (e.g., brick, concrete, wood) Brick and Mortar
    # # Wall Thickness (in meters): 0.20
    # # Number of Doors: 1
    # # Door Dimensions (width x height x thickness, in meters): 2 x 1 x 0.20
    # # Type of Door (e.g., wooden, metal, glass): Wooden
    # # Number of Windows (if any): 2
    # # Window Dimensions (width x height, in meters) 1 x 1
    # # Type of Window (e.g., sliding, casement, double-hung): Fixed
    # # Paint Type and Color: Flat and White
    # # Paint Thickness (in millimeters): 2
    # # Ceiling and Roof Required (Yes/No): Yes
    # # Electrical Work Required (Yes/No): No
    # # Plumbing Work Required (Yes/No): No
    # # Project Deadline (in weeks): None"""
    gpt_reponse = get_response_from_ChatGPT(user_prompt)
    # print(gpt_reponse)
#     gpt_reponse = """[['Task name', 'Duration', 'Start date', 'End date', 'Predecessor'],
# ['Site Preparation', 1, 'Jul 26, 2023', 'Jul 27, 2023', ''],
# ['Foundation Construction', 5, 'Jul 28, 2023', 'Aug 1, 2023', '1'],
# ['Wall Construction', 5, 'Aug 2, 2023', 'Aug 6, 2023', '2'],
# ['Door Installation', 1, 'Aug 7, 2023', 'Aug 8, 2023', '3'],
# ['Window Installation', 1, 'Aug 7, 2023', 'Aug 8, 2023', '3'],
# ['Painting', 2, 'Aug 9, 2023', 'Aug 10, 2023', '4,5'],
# ['Ceiling and Roof Construction', 3, 'Aug 11, 2023', 'Aug 13, 2023', '6'],
# ['Project Completion', 1, 'Aug 14, 2023', 'Aug 15, 2023', '7']]"""
    create_mpp_file(gpt_reponse)
