import os
import json
import azure.functions as func
import openai

def main(req: func.HttpRequest) -> func.HttpResponse:
    prompt = req.params.get('prompt', 'Hello')
    engine = req.params.get('engine', "text-davinci-003")
    temperature = float(req.params.get('temperature', 0.7))
    max_tokens = int(req.params.get('max_tokens', 30))
    verbose = req.params.get('verbose', False)

    openai.api_type = "azure"
    openai.api_key = os.getenv("AZURE_OPENAI_KEY")
    openai.api_base = os.getenv("AZURE_OPENAI_ENDPOINT")
    openai.api_version = "2023-03-15-preview" 

    response_json = openai.Completion.create(engine=engine, prompt=prompt, temperature=temperature, max_tokens=max_tokens)

    response_text = response_json['choices'][0]['text']  #parse text (prompt completion)

    if verbose == True:
        response = {
            "prompt": prompt,
            "response_text": response_text
        }
    else:
        response = {
            "response_text": response_text
        }

    #return func.HttpResponse(body=json.dumps(response), mimetype="application/json")
    return func.HttpResponse(body=response_text, mimetype="text/plain")