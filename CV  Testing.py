import os
from pathlib import Path
import datetime
import re
import json
import win32com.client
from win32com.client import Dispatch
import PyPDF2
from openai import OpenAI
from termcolor import colored

# Replace with your API key, model, and messages
api_key = "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"
model_name = "mistral-7b-instruct"


def download_attachments_from_outlook(date_):
    try:

        # Connect to Outlook
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

        # inbox = outlook.Folders["project@my-mj.de"].Folders["newProject"].Folders["Posteingang"]
        inbox = outlook.GetDefaultFolder(6)

        # Get messages
        messages = inbox.Items

        for message in messages:
            subject = message.Subject
            received_time = message.ReceivedTime
            attachments = message.Attachments

            # Check if email subject contains "CV" and received on or after the specified date
            if ("cv" in subject.lower() or "bewerbung" in subject.lower() or "jobds" in subject.lower()) and received_time.date() >= date_.date():

                pdf_attachments = [attachment.FileName for attachment in attachments if
                                   attachment.FileName.endswith(".pdf")]

                if len(pdf_attachments) == 1:  # Check if there's only one attachment

                    for attachment in attachments:
                        if attachment.FileName.endswith(".pdf"):
                            try:
                                folder_name = received_time.strftime("%Y-%m-%d")  # Format timestamp
                                folder_path_ = os.path.join("C:/Users/DELL/PycharmProjects/CV/Attachments", folder_name)

                                # Create folder if it doesn't exist
                                if not os.path.exists(folder_path_):
                                    os.makedirs(folder_path_)

                                file_path = os.path.join(folder_path_, attachment.FileName)

                                # Check if file already exists, if yes, skip saving
                                if os.path.exists(file_path):
                                    print(f"Attachment '{attachment.FileName}' already exists. Skipping.")
                                else:
                                    attachment.SaveAsFile(file_path)
                                    print(f"Attachment saved: {file_path}")
                            except Exception as e:
                                print(f"Error saving attachment '{pdf_attachments[0]}': {e}")


                elif len(pdf_attachments) > 1:  # If there are multiple attachments, loop over each one
                    for attachment in attachments:
                        if attachment.FileName.endswith(".pdf"):
                            try:
                                if "cv" in attachment.FileName.lower() or "resume" in attachment.FileName.lower() or "lebenslauf" in attachment.FileName.lower():
                                    folder_name = received_time.strftime("%Y-%m-%d")  # Format timestamp
                                    folder_path_ = os.path.join("C:/Users/DELL/PycharmProjects/CV/Attachments",
                                                                folder_name)

                                    # Create folder if it doesn't exist
                                    if not os.path.exists(folder_path_):
                                        os.makedirs(folder_path_)

                                    file_path = os.path.join(folder_path_, attachment.FileName)

                                    # Check if file already exists, if yes, skip saving
                                    if os.path.exists(file_path):
                                        print(f"Attachment '{attachment.FileName}' already exists. Skipping.")
                                    else:
                                        attachment.SaveAsFile(file_path)
                                        print(f"Attachment saved: {file_path}")
                            except Exception as e:
                                print(f"Error saving attachment '{attachment.FileName}': {e}")


    except Exception as e:
        print(f"Error downloading attachments from Outlook: {e}")


def do(date_, folder_path_):
    download_attachments_from_outlook(date_)
    dict_pdf = extract_text_from_pdf(folder_path_, date_)

    for k in dict_pdf:
        # Reset conversation messages for each PDF
        conversation_messages = [
            {"role": "system", "content": "Your system message here"},
            {"role": "user", "content": """Instructions:

        1. Analyze the CV PDF to extract the following information:
            * Name: Full name of the candidate.
            * Location: Residential address, exclude work address (Pick that Address where person currently living), comes after person name usually.
            * Language: Proficiency level in various languages, focusing on German (B1, B2, C1, C2).
            * Experience: Total years of relevant experience.
            * Job Background: Summarize relevant job experience, focusing on companies, positions, and key responsibilities.
        2. Apply the following rating system:
            * AAA: Candidate resides in Germany or India, and has B2 and higher German proficiency, possesses relevant keywords like Banking, IT, and Finance, and has any level of experience.
            * BBB: Candidate resides in Germany or India, and has B1, B2 or higher German proficiency, possesses relevant keywords like Banking, IT, and Finance, and has less than or equal to 4 years of experience or exhibits frequent job changes.
            * CCC: Candidate resides outside of Germany or India or does not possess German language skills.
        3. Keywords:
            * Banking: Compliance, Payments, Transactions, Cards, AML, AFC, KYC, Regulatory Reporting, Banking, Finance, Insurance, BAFIN, MIFID.
            * IT: BPMN, UML, QA testing, test, Business Analysis, Analyst, IT Consultant, PMO, Project Manager, Data Analyst, Application Manager, DORA, CSRD, SCRUM, Audit, namelist screening, Embargo, Provider Management, IT Security, Cyber Security, network Security, LAN/WAN ISO27001, SWIFT, SEPA, AWS, BI.
        4. Note:
            * If the candidate has worked in multiple locations, sum their total experience and provide an overall summary of their job background.
            * Ignore the working address when extracting the location information.
        5. Give information as per requested and Rate them according requirement AAA, BBB or CCC not netural or something 
        6. Candidate resides outside of Germany or India, Give rating "CCC" straight away 
        7. Give in This format Name, Location, Language Level, Job Background, Rating    : """},
        ]

        conversation_messages_1 = [
            {"role": "system", "content": "Your system message here"},
            {"role": "user", "content": """find any information on internet,public and social platforms, 
                positive and negative and then rate if the candidate is clean from any legal or fraud trouble
                 regarding followed mention person
                  Give reference links regarding information : """},
        ]

        # Get text from PDF
        pdf_text = dict_pdf[k]

        # Append PDF text to conversation messages
        conversation_messages[-1]["content"] += f"\n\n{pdf_text}"

        # Call OpenAI for CV rating
        result = chat_with_openai(api_key, model_name, conversation_messages)

        # Extract name and rating
        name = extract_name(result)
        rating_ = rating(result)

        # Append name check prompt to conversation messages
        conversation_messages_1[-1]["content"] += f"\n\n{name}"

        # Call OpenAI for name check
        result_1 = chat_with_openai(api_key, model_name, conversation_messages_1)

        a = "DISCLAIMER: this is an AI generated information and can be untrue, please handle with caution and double check the details"


        # Prepare email subject and body
        subject = f"JOBDS AI Response:- {name}: CV Rating >> {rating_}"
        body = f"{a}\n\nAssistant's Response:\n\n{result}\n\n\n\nName Check:\n\n{result_1}"

        send_email(subject, body, k)       # "k" : pdf path for an attachment



def extract_text_from_pdf(folder_path_, date_):
    pdf_dict = {}
    for folder in os.listdir(folder_path_):

        sub_folder_path = os.path.join(folder_path_, folder)
        if os.path.isdir(sub_folder_path):
            folder_date = datetime.datetime.strptime(folder, "%Y-%m-%d").date()
            if folder_date >= date_.date():
                print(f"Sub folder: {folder}")

                for file in os.listdir(sub_folder_path):
                    if file.lower().endswith('.pdf'):

                        try:
                            with open(os.path.join(sub_folder_path, file), 'rb') as pdf_file:
                                pdf_reader = PyPDF2.PdfReader(pdf_file)
                                text = ""
                                for page_num in range(len(pdf_reader.pages)):
                                    page = pdf_reader.pages[page_num]
                                    text += page.extract_text()
                                pdf_dict[os.path.join(sub_folder_path, file)] = text

                        except Exception as e:
                            print(f"Error extracting text from PDF: {e}")

    return pdf_dict


def chat_with_openai(api_key_, model, messages):
    try:
        # Access the OpenAI API (modify base_url for Perplexity.ai)
        client = OpenAI(api_key=api_key_, base_url="https://api.perplexity.ai")

        # Send the chat completion request
        response = client.chat.completions.create(
            model=model,
            messages=messages,
        )

        # Serialize the relevant information to a JSON string
        response_json_str = json.dumps({
            "choices": [
                {
                    "finish_reason": response.choices[0].finish_reason,
                    "index": response.choices[0].index,
                    "message": {
                        "content": response.choices[0].message.content,
                        "role": response.choices[0].message.role,
                    },
                    "delta": response.choices[0].delta,
                }
            ],
            "created": response.created,
            "model": response.model,
            "object": response.object,
            "usage": {
                "completion_tokens": response.usage.completion_tokens,
                "prompt_tokens": response.usage.prompt_tokens,
                "total_tokens": response.usage.total_tokens,
            }
        })

        # Convert the serialized JSON string to a dictionary
        response_dict = json.loads(response_json_str)

        # Extract and return the assistant's response
        assistant_response = response_dict["choices"][0]["message"]["content"]
        return assistant_response

    except Exception as e:
        print(f"Error chatting with OpenAI: {e}")
        return ""


# Extract Name

def extract_name(text):
    """Extracts the name from a given text string.

  Args:
    text: The text string to extract the name from.

  Returns:
    The extracted name if found, otherwise None.
  """

    # Combine patterns to handle various name formats, excluding location:
    name_patterns = [
        r"Name:\s+([\w\s]+)(?:\s+|$)",
        # "Name:" followed by space and one or more words/spaces (non-greedy match to avoid location)
        r"([\w\s]+?)\s+(?!Location:)(?:Job Background|Experience):",
        # Name followed by whitespace and "Job Background" or "Experience" (excluding "Location:", handles initials and middle names)
        r"(\w+)\s+(\w+)\s+(?!Location:)(?:Job Background|Experience):",
        # First, middle, and last name separately (for cases like Yuliia Misnichenko)
    ]

    for pattern in name_patterns:
        match = re.search(pattern, text)
        if match:
            # Handle different match group arrangements depending on the pattern:
            if len(match.groups()) == 1:
                return match.group(1).strip()
            elif len(match.groups()) == 2:
                return " ".join(match.groups()).strip()
            else:
                return " ".join(match.groups()[1:]).strip()  # Skip first group (first name)

    # No match found
    return None


def rating(text):
    # Define a regex pattern to match any combination of three letters
    pattern = r"\b[A-C]{3}\b"

    # Search for the pattern in the text
    match = re.search(pattern, text)

    # If a match is found, return the rating, otherwise return None
    return match.group(0) if match else None


def send_email(subject, body, attachment_path):
    """
  Sends an email through Outlook.

  Args:
    :param subject:
    :param body:
    :param attachment_path:
  """

    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)

    mail.To = "work@mejuvante.net "          #"ankita.rishabh@mejuvante.co.in"
    mail.Subject = subject
    mail.Body = body

    mail.Attachments.Add(attachment_path)

    mail.Send()
    # outlook.Quit()


# Convert date string to datetime object
date_str = input("Enter the date (YYYY-MM-DD): ")
date = datetime.datetime.strptime(date_str, "%Y-%m-%d")
folder_path = r"C:\Users\DELL\PycharmProjects\CV\Attachments"
do(date, folder_path)
