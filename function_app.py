import azure.functions as func
import logging
import requests
import json
import pandas as pd
from msal import ConfidentialClientApplication
import openai
from io import StringIO

# Replace with your Graph API credentials and endpoints
GRAPH_API_ENDPOINT = "https://graph.microsoft.com/v1.0"
access_token = "EwCIA8l6BAAUbDba3x2OMJElkF7gJ4z/VbCPEz0AAZmhOCrb7WZ/z44aTxSmGJhVPMYzVKxRZJ270rFGyQSlBpx5JDS+rf68YbHslaaa8wtod/tyNTk56H6BDP7hdydWPfPYpkNyx0qctwBM/ZOtWYHmD1KEoALJlBo54GEzrJGyrCWAyob1q6Qt06DrEVYaJDz0SNPLdX+Hv0nLMt2kw6cIxfCH9m0fsuH8oYwkYvtHhSXHNJQnPc7DxdjUvT7WqLUC0qvnc7nHxVEfoQdaM6R60TJXH5azLFbkITdERgiJw9/Rdc0mMZQzHDjp+bFjGklXQMSBbci7muHmlsOiwp7KaSoJpHFbC/eToqiSMs3Gk2uwsym2yGR+D2Tagl0QZgAAEDeJiFS03KOrQmtRQlTHqGZQAofY7cOCUpFvTj89KvsZOJJ4514J7eWGPMRDDoRDpOYOSuBCn5Sj+XeI5qQ+eyKCOeO4D/iA8t+AeznUx9cbnaUYvGqkBB4dpI1Ehpe9eGuChwJGPTlBNBmfn7lhRmjkWK2rekaR5zbW52G3lqljo7M06su/ojzt1R082s1K6klgE8el3HvPwKzb/DcylXzXT0AoCIjRAXmbK4T4P7cP5EXVTFl6NS44qvnsoN3npk4qboPFDcSKfDYkjaFAk1XZOacP/Kk4LoFRGYo3nzCvF9SW/dnuRg4dsMGGNopyKZn9DVE7Cl8n9jtqys7jwlthURM4R6/Y13jJVB1KoBVizMdFW0u7I7Rm7wQu6y+3wJJXnIErMxtUywagk3O6FRZzguZfoQv7UBCKUzcfrRHXUmK84roJ52pGMgHhmMCzbnq0AkkRzOmD0YIZlTxSXzVA707B8pAbrTdadiDm2BzYJQ9Vr0zX4KcYsK/xhlvh1l0KdwlkzmMWmU06yk2R1qP514BNArjU6sk+stoRCqJ/7CHEIHiYrrzYGTu2UlBsk7tZrvBYrvbwGkpNY3bspk5R7gLzYk8QHL8d0lHNIT5Pz3QBS10akaAOe8zNV0cmf0H5CF53G2KkHOB4Mzs/1OpMyxB/trB7R9UnrsrE3xQFdNfpuwKhhMP7UjTEjd2SK0pJN5OC1AIYlslk7JffiYs+MKJkVRksfwqpxGGyzVK2kiivswywoDFnpaS6Lq9QzfWeegFrYv6c9LaSU5ZvphUV3RZCNzInlQRiEZZBeE44r9+OAg=="  # Replace with actual token retrieval logic

openai.api_key = "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"

app = func.FunctionApp(http_auth_level=func.AuthLevel.FUNCTION)

@app.route(route="EmailProcessor")
def EmailProcessor(req: func.HttpRequest) -> func.HttpResponse:
    logging.info('Python HTTP trigger function processed a request.')

    # Step 1: Set up MSAL (Microsoft Authentication Library)
    client_id = '792aede1-cbeb-4709-b1b4-dac0badcd91c'
    client_secret = 'JsQ8Q~b5RhBoazz2~3IhjBFOBcxW3ZUXLV_qVb3d'
    tenant_id = '242619cc-5895-4d94-9475-aec1e3de9b59'
    authority = f"https://login.microsoftonline.com/{tenant_id}"
    scope = ["https://graph.microsoft.com/.default"]

    # Retrieve and process the request parameters
    mailbox = req.params.get('mailbox')
    if not mailbox:
        return func.HttpResponse(
            "Please pass the 'mailbox' parameter in the query string.",
            status_code=400
        )

    # Step 2: Check for new emails
    headers = {"Authorization": f"Bearer {access_token}"}
    email_endpoint = f"{GRAPH_API_ENDPOINT}/users/karanoza586@outlook.com/messages"
    
    try:
        response = requests.get(email_endpoint, headers=headers)
        response.raise_for_status()  # Raise an error for HTTP errors
    except requests.exceptions.HTTPError as http_err:
        logging.error(f"HTTP error occurred: {http_err}")
        return func.HttpResponse(
            f"HTTP error occurred: {http_err}",
            status_code=500
        )
    except Exception as err:
        logging.error(f"Other error occurred: {err}")
        return func.HttpResponse(
            f"Other error occurred: {err}",
            status_code=500
        )

    messages = response.json().get('value', [])
    if not messages:
        return func.HttpResponse("No new emails found.", status_code=200)

    # Step 3: Process the emails and save subjects to a CSV
    email_data = []
    for message in messages:
        email_subject = message['subject']
        
        # OpenAI Integration: Generate a response for each email subject
        prompt = f"Analyze the following email subject and summarize its key point: '{email_subject}'"
        openai_response = openai.ChatCompletion.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "system", "content": "You are an email summarizer."},
                {"role": "user", "content": prompt}
            ]
        )
        
        generated_text = openai_response.choices[0].message['content'].strip()
        email_data.append({'Subject': email_subject, 'AI Response': generated_text})


    # Convert email subjects to CSV format
    df = pd.DataFrame(email_data)
    print(df)
    csv_buffer = StringIO()
    df.to_csv(csv_buffer, index=False)

    # Step 4: Upload CSV to OneDrive
    upload_to_onedrive(csv_buffer.getvalue(), headers)

    return func.HttpResponse("Emails processed and saved successfully!", status_code=200)

def upload_to_onedrive(csv_content, headers):
    # Assuming OneDrive integration here (simplified example)
    upload_endpoint = f"{GRAPH_API_ENDPOINT}/me/drive/root:/EmailData.csv:/content"
    upload_response = requests.put(upload_endpoint, headers=headers, data=csv_content)
    logging.info(f"Upload response: {upload_response.status_code} - {upload_response.text}")



@app.route(route="EmailProcessor", auth_level=func.AuthLevel.FUNCTION)
def EmailProcessor(req: func.HttpRequest) -> func.HttpResponse:
    logging.info('Python HTTP trigger function processed a request.')

    name = req.params.get('name')
    if not name:
        try:
            req_body = req.get_json()
        except ValueError:
            pass
        else:
            name = req_body.get('name')

    if name:
        return func.HttpResponse(f"Hello, {name}. This HTTP triggered function executed successfully.")
    else:
        return func.HttpResponse(
             "This HTTP triggered function executed successfully. Pass a name in the query string or in the request body for a personalized response.",
             status_code=200
        )