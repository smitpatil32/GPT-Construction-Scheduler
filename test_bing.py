# import BingPython as ai
import asyncio
import re
# import win32com.client
import datetime
import os

api_key = os.get("CHAT_GPT_API_KEY")

# cookies = {
#         'SnrOvr':	'X=rebateson',
#         'SRCHUSR':	'DOB=20230413&T=1683241943000&POEX=W',
#         'ai_session':	'ykfy8+I1K/UFB7n6FnM7HR|1683241944431|1683242041489',
#         'SUID':	'A',
#         'SRCHHPGUSR':	'SRCHLANG=en&DM=0&WTS=63818830893&PV=5.15.0&BRW=XW&BRH=M&CW=1791&CH=980&SCW=1776&SCH=3331&DPR=1.0&UTC=-240&EXLTT=9&HV=1683242062&PRVCW=1791&PRVCH=980',
#         'ANON':	'A=26640780E28C8421F4A053DEFFFFFFFF&E=1c36&W=1',
#         'MicrosoftApplicationsTelemetryDeviceId':	'b6fb9518-ca8c-4347-b635-b5b16ceb3521',
#         '_SS':	'SID=27626A415B7566B8159279465A0B6766&R=15&RB=15&GB=0&RG=0&RP=15',
#         'ipv6':	'hit=1683245545329&t=4',
#         'dsc':	'order=News',
#         '_U':	'10Ncy3NtF_MUjTRTWUbmjC9hLiAxb4TxoR46nsphYakLRqLD4O9TyjYrEixri7aybVNgw0y9T43zLyRs-F3288q2jDqH75LvTZsHxS_vwJN0npuHJ57ypIn3gTTnMtIs_9gaTmyXmybBdkwBO3hYiPLFg5LLUcqBpbdhlT_OEOEKjLHJBaTatw5KQzLVVKq4n947sw0qcisufzY0KTIgnG-mnVEblb0Z72m8XU8O3BoI',
#         'SRCHD':	'AF=SHORUN',
#         'PPLState':	'1',
#         'NAP':	'V=1.9&E=1bdc&C=csz9B6Dq4BHV7zr9Bwhm088lexcTbaBGlOOhi1k2Z1zZhMtLWTAtPg&W=1',
#         '_RwBf':	'ilt=2&ihpd=0&ispd=1&rc=15&rb=15&gb=0&rg=0&pc=15&mtu=0&rbb=0.0&g=0&cid=&clo=0&v=10&l=2023-05-04T07:00:00.0000000Z&lft=0001-01-01T00:00:00.0000000&aof=0&o=0&p=bingcopilotwaitlist&c=MY00IA&t=9513&s=2023-04-13T18:11:23.5471026+00:00&ts=2023-05-04T23:14:19.7164748+00:00&rwred=0&wls=2&lka=0&lkt=0&TH=&r=1&mta=0&e=IRJEjfgkLBnWJo2cYz-krvCEjWDU6-zMCb-JYcRdKDp4HH2JVrVe2haOF0MmaoRn44U9rdEEAUy_UdP9yZtFPA&A=26640780E28C8421F4A053DEFFFFFFFF',
#         '_UR':	'QS=0&TQS=0',
#         '_EDGE_S':	'SID=27626A415B7566B8159279465A0B6766&mkt=en-us',
#         'MUIDB':	'0816A2708D0266CE2986B0828CD167B3',
#         'USRLOC':	'HS=1&ELOC=LAT=40.42902755737305|LON=-86.90729522705078|N=West%20Lafayette%2C%20Indiana|ELT=4|',
#         '_clck':	'1napbi2|1|faq|0',
#         '_HPVN':	'CS=eyJQbiI6eyJDbiI6MSwiU3QiOjAsIlFzIjowLCJQcm9kIjoiUCJ9LCJTYyI6eyJDbiI6MSwiU3QiOjAsIlFzIjowLCJQcm9kIjoiSCJ9LCJReiI6eyJDbiI6MSwiU3QiOjAsIlFzIjowLCJQcm9kIjoiVCJ9LCJBcCI6dHJ1ZSwiTXV0ZSI6dHJ1ZSwiTGFkIjoiMjAyMy0wNS0wNFQwMDowMDowMFoiLCJJb3RkIjowLCJHd2IiOjAsIkRmdCI6bnVsbCwiTXZzIjowLCJGbHQiOjAsIkltcCI6NX0=',
#         'KievRPSSecAuth':	'FACSBBRaTOJILtFsMkpLVWSG6AN6C/svRwNmAAAEgAAACFH0CxkAYuXFUATHbWZAHtBRulSY7xG4QlYkLp5cDYTujXl9LAGkRC4ev9IG/l5kukwOHgztxQtDEpAYATC4VAiIs6eyUrHkLqzepLECjJTpxARbznp+0lVqiHucDcvSvG9xVCj5ldG+Aoc3CCXQOb9mdEvfUzEZO0kHAIIFurwmX7oN92v3kJfzpn5/RhfYJmPCtZ0iF3W6oDItoeT+m+1DnWaOIl8keMLAHIR/BXUlS0ZKV5OlceCYGRRgMwJ9zB5c4boZqS0Ttc73vT+1IbfDc/scq4qJtVeSH9xhre5CgCxg/z/7Xy8BytRD0VGhJD+dzXcfiS7+foR0iGCnoYNU1VX/2ru1tcYkQTvS+C8Y/zNK3eJahSFI44j5TkAec9RCcMLyJkPeEPne7Jh6F87N8lX0OSGk4itaeo0jlhfu/ouVRb2S/jnrLLZ4EXhUWsqiOyd6iEOUVRQsuqlKnHxijqHO0jHofhZ85JpkmTsUTieYSF+c9UOcwejJ/3/gjujdbkqClwowsGW/F+96OhoVuBWMI/YhhrU/3b1js7cy7XRKGJeRxVBAci27rsqmbik5BKcZdiCOhYVfGi6txkhLopa49gI/lSaf2brwq8Q0D9MCNfkNTJTCIWraPFVIXQ3S4o0gm6VhOpkT7GqI108X2WfR+CDZmjKB/RmpJLe/+MFcYYDSO5rc9AraRc7h9Sfpm6Pf4TnuIzxseijF1S6sCKcUnxfPeuTwXoiphImBuGmZvHieOA6MQNAswyga3zV1sgA3+6odMJ7OJ5L4wr/oghWa5I3wo5OKD47W4BOFxydOsazA2grCYpMw3I3z2DcvU1O7cqrfVa5I/5JT2OJ5Gxql0rGNq7giy7H1gQktpxTtsxksZEIFmwGLtIVlq4nH8Gb+9M+eX/DuH5vkfcKxte+b8MGHQzulSdYRN1fPZsgZOMvg1c2eazrkzZYl7G1poeBDUOleFk5g71nBtfOlyFf3jaKJBuohva5zujogp73eXYNZ/1UH5POoGGJkzIzXAQ6pRjzlu+c6L6jQypSA6iN0Fabv5vvjyeOxfB2Fha4aSNvJz8Tn794iwiXRQ4cnoXcwDWEb9W/grKKmjwEkLMnf8taav+8eWgsucU4BF8MoTBVHP1ul9oYX7B4k2g8WZzsRyZO8LqEDzkl5uoifCMPx3dtqvAkDy03YPGPxXJ5nFoGpdp+bW2t2Xo2p/5bt+4ntb6s5J3X9yIwVauFDN1Vs5jMSpY9tAsdKpAJhVGIKOLUP2ZHSRyK5ONpDtB0i5m85a/XVL1Y1GOeldFdvk+WGnjnwvaFMKLEhHqMLQIMrBIYEQBoq8HWJa91W531gPaFymgs2VbAsG/HyBO6V0zL2TBVXq1jsv+skJlH+XV1tiatYGcD/V12qlDCLMIVlch6e7T09Gp7YVxvztY6kkRhsX2DuUJGTs2l3B1KOExIO0/fmBGxApqvtSilvELGime0mdnVn8+AUANd0gjKAXjwJBW6B0qdVtGFhebD4',
#         'msaoptout':	'0',
#         'MUID':	'0816A2708D0266CE2986B0828CD167B3',
#         'SRCHUID':	'V=2&GUID=F418736368F24F2CA8D0E3D99CBE0A88&dmnchg=1',
#         'WLID':	'lTDJ9rIrOwV79N69MO4R153dk6lImZPkJShMnAsHzFy6aAFd3FKBQvDARAB8XIEUm7P0GSwvnNL76omADxPG6++4pdracFmKWeE8VBfwsGI=',
#         'WLS':	'C=c3fa2fa4dfa5654d&N=trivikram'
#     }

def extract_cookies():
    cookies = {}
    with open("cookie.txt", "r") as f:
        for line in f:
            line = line.strip().split("\t")
            if len(line) > 2:
                key = line[-2]
                value = line[-1]
                cookies[key] = value
    return cookies

def get_details_from_user():

    details_str = ""
    prjname = input("Project Name: ")
    details_str += f"Project Name: {prjname}; "
    location = input("Project Location: ")
    details_str += f"Project Location: {location}; "
    projtype = input("Project Type (e.g., residential, commercial, industrial): ")
    details_str += f"Project Type: {projtype}; "
    dimensions = input("Dimensions (length x width x height, in meters): ")
    details_str += f"Dimensions (length x width x height, in meters): {dimensions}; "
    material = input("Type of Construction Material (e.g., brick, concrete, wood): ")
    details_str += f"Type of Construction Material (e.g., brick, concrete, wood): {material}; "
    wallthickness = input("Wall Thickness (in meters): ")
    details_str += f"Wall Thickness (in meters): {wallthickness}; "
    numdoors = input("Number of Doors: ")
    details_str += f"Number of Doors: {numdoors}; "
    doordim = input("Door Dimensions (width x height x thickness, in meters): ")
    details_str += f"Door Dimensions (width x height x thickness, in meters): {doordim}; "
    doortype = input("Type of Door (e.g., wooden, metal, glass): ")
    details_str += f"Type of Door (e.g., wooden, metal, glass): {doortype}; "
    numwindows = input("Number of Windows (if any): ")
    details_str += f"Number of Windows: {numwindows}; "
    windowdim = input("Window Dimensions (width x height, in meters): ")
    details_str += f"Window Dimensions (width x height, in meters): {windowdim}; "
    windowtype = input("Type of Window (e.g., sliding, casement, double-hung): ")
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

    return details_str

if __name__ == "__main__":
    
    query_string = "With the following 'construction details information' create a schedule for the construction project described. The schedule should include a start date and an end date and breakdown the construction project into specific tasks and mention the time allocated for each task in days. Your response should only represent the schedule in a tabular form. The table should contain following columns in this specific order - Task name, Duration, Start date, end date. As an example: Site Preparation | 2 | Mar 23, 2020 | Mar 25, 2020|. I do not want any citations or explanations. "
    query_string += get_details_from_user()
    # query_string += "Project Name: Sample project; Project Location: Indiana; Project Type (e.g., residential, commercial, industrial): Residential; Dimensions (length x width x height, in meters): 4.2 x 4.2 x 4.2; Type of Construction Material (e.g., brick, concrete, wood): Brick Masonary; Wall Thickness (in meters): 0.20; Number of Doors: 1; Door Dimensions (width x height x thickness, in meters): 1 x 2 x 0.20; Type of Door (e.g., wooden, metal, glass): Wooden; Number of Windows (if any): 2; Window Dimensions (width x height, in meters):  2 x 1 ; Type of Window (e.g., sliding, casement, double-hung): Fixed Window; Paint Type and Color: Flat Paint White; Paint Thickness (in millimeters): 10 ; Ceiling and Roof Required (Yes/No): Yes; Electrical Work Required (Yes/No): No; Plumbing Work Required (Yes/No): No; Project Deadline (in weeks): None"
    print(query_string)

    cookies = extract_cookies()

    # command = ai.BingPython.sendcom_sydney(ai.BingPython.CreateSession(cookies), query_string)
    # result = asyncio.get_event_loop().run_until_complete(command)
    # print(result)

    # Index, Taskname, duration, startdate, enddate
    matcher1 = re.compile(r'\| ([0-9]+) \| ([A-Za-z\s]+) \| ([0-9]+) \| ([A-Za-z0-9,\s]+) \| ([A-Za-z0-9,\s]+)')
    # Taskname, duration, startdate, enddate
    matcher2 = re.compile(r'\| ([A-Za-z\s]+) \| ([0-9]+) \| ([A-Za-z0-9,\s]+) \| ([A-Za-z0-9,\s]+)')

    tasks = []
    result_lines = result.split("\n")
    for line in result_lines:
        matches = matcher1.findall(line)
        if not matches:
            matches = matcher2.findall(line)
        if matches:
            matches = matches[0]
            if len(matches) == 5:
                index = matches[0]
                task_name = matches[1]
                duration = matches[2]
                start_date = matches[3]
                end_date = matches[4]
            elif len(matches) == 4:
                task_name = matches[0]
                duration = matches[1]
                start_date = matches[2]
                end_date = matches[3]
            tasks.append({"name": task_name, "duration": duration})

    print(tasks)

    # Connect to the Microsoft Project application
    app = win32com.client.Dispatch("MSProject.Application")
    setattr(app, "Visible", True)  # Make the application window visible

    # Create a new project
    project_name = "Construction Project"
    new_project = app.Projects.Add()
    new_project.Name = project_name

    # Add the tasks to the project
    start_date = datetime.datetime.now()
    for i, task in enumerate(tasks):
        task_name = task["name"]
        duration = task["duration"] #* 1440  # Convert duration to minutes (1 day = 1440 minutes)
        new_task = new_project.Tasks.Add(task_name)
        new_task.Duration = duration
        new_task.Manual = False  # Set the task as auto-scheduled
        if i == 0:
            new_task.Start = start_date
        else:
            new_task.Predecessors = str(new_project.Tasks(i).ID)

    # Check if the folder exists and create it if necessary
    folder_path = "D:\My Files\Courses\Spring 2023\Automation in Construction\Framework for Automation\Trivi" #"D:\\sad\\"
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)

    # Save the project
    file_path = f"D:\My Files\Courses\Spring 2023\Automation in Construction\Framework for Automation\Trivi\hello.mpp"
    new_project.SaveAs(file_path)

    # Open the newly created project
    app.FileOpen(file_path)

    # Close the project using app.FileClose()
    app.FileClose()
