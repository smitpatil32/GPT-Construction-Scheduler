import openai
import os
import ast
import win32com.client
import datetime

openai.api_key = os.getenv("CHAT_GPT_API_KEY")


def get_details_from_user():
    main_string = "With the following 'construction details information' create a schedule for the construction project described. The schedule should include a start date and an end date and breakdown the construction project into specific tasks and mention the time allocated for each task in days. Your response should only represent the schedule in a tabular form. The table should contain following columns in this specific order - Task name, Duration, Start date, end date. As an example: Site Preparation | 2 | Mar 23, 2020 | Mar 25, 2020|. I do not want any citations or explanations. "
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
    response = openai.ChatCompletion.create(
        model="gpt-3.5-turbo",
        messages=[
            {"role": "system", "content": "You are an expert construciton project scheduler."},
            {"role": "user", "content": prompt},
        ]
    )
    return response.choices[0].message.content


def create_mpp_file(response_string):
    # Create a list from the given chatgpt output
    tasks_list = ast.literal_eval(response_string)

    project_name = input("Enter the Project Name: ")
    project_name += "_"+datetime.datetime.now().strftime('%m-%d-%y_%H-%M-%p')

    # Initialize the MS Project App
    Project_App = win32com.client.Dispatch("MSProject.Application")
    Project_App.Visible = True
    pj = Project_App.Projects.Add()

    for task in tasks_list[1:]:
        startdate = datetime.datetime.strptime(task[2], "%b %d, %Y")
        enddate = datetime.datetime.strptime(task[3], "%b %d, %Y")
        duration = task[1]
        task_adder = pj.Tasks.Add(task[0])
        task_adder.Duration = duration
        task_adder.Start = startdate
        task_adder.Finish = enddate

    pj.SaveAs(os.path.join(os.getcwd(), f"{project_name}.mpp"))
    Project_App.Quit()


# test_output = """[['Task name', 'Duration', 'Start date', 'End date'],
# ['Site Preparation', 2, 'Jul 1, 2021', 'Jul 2, 2021'],
# ['Foundation', 7, 'Jul 5, 2021', 'Jul 13, 2021'],
# ['Brickwork', 14, 'Jul 14, 2021', 'Jul 30, 2021'],
# ['Door Installation', 2, 'Aug 2, 2021', 'Aug 3, 2021'],
# ['Window Installation', 2, 'Aug 4, 2021', 'Aug 5, 2021'],
# ['Painting', 5, 'Aug 6, 2021', 'Aug 12, 2021']]"""

if __name__ == '__main__':
    user_prompt = get_details_from_user()
    gpt_reponse = get_response_from_ChatGPT(user_prompt)

    create_mpp_file(gpt_reponse)
