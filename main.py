from chat_gpt import get_response_from_ChatGPT, get_details_from_user, create_mpp_file
# from bing import main_bing

AI_CHOICE = input(
    "\nPlease choose the AI of your choice [C] for ChatGPT [B] for Bing AI C/B: ")


def start(AI):
    if AI.lower() == "c":
        
        user_prompt = get_details_from_user()
        gpt_reponse = get_response_from_ChatGPT(user_prompt)
        create_mpp_file(gpt_reponse)
    # elif AI.lower() == "b":
    #     main_bing()
    else:
        return "WRONG INPUT TRY AGAIN."


if __name__ == '__main__':
    start(AI_CHOICE)
