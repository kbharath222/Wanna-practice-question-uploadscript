import json
import pandas as pd
import requests
from certificateMapping import certificate_mapping
from itertools import islice
from login import login, baseUrl, signout
from docxconverter import read_docx,write_excel
from datetime import date


'''
1) take the input from user and compare with certification header (done)
2) check for the existing question
3) save the payload with proper results with date time
'''

url = baseUrl()
params = login()
question_count = 0

if params: # if login is successful then execute the code
    
    certification_choice = input('Enter the choice for certification name: 1) for CISA, 2) for CISSP, 3) for CCSP, 4) for SSCP, 5) for CISM, 6) for CISSP-2024 :') #taking input from user

    cert_mapping = {  # Mapping of certification based on user choice
        "1": "CISA",
        "2": "CISSP",
        "3": "CCSP",
        "4": "SSCP",
        "5": "CISM",
        "6": "CISSP-2024"
    }

    cert_name = cert_mapping.get(certification_choice).upper()

    certification_choice_mapping_with_column_name = { # Mapping of certification with excel column
        "1": "CISA",
        "2": 'CISSP',
        "3": "CCSP",
        "4": "SSCP",
        "5": "CISM",
        "6": "CISSP-2024"
    }

    cert_column_name = certification_choice_mapping_with_column_name.get(certification_choice)

    # converting docx file to excel
    input_docx_path='batch_34.docx'
    doc_data=read_docx(input_docx_path)
    today= date.today()
    output_name = f'{input_docx_path.split(".")[0]}, {today}'
    output_excel_path = f'excelsheets/{output_name}.xlsx'
    write_excel(doc_data, output_excel_path)
    

    path = output_excel_path      # path where the excel sheet is located

    data = pd.read_excel(path, sheet_name="CISSP", usecols=['Question', 'Option A', 'Option B', 'Option C', 'Option D', 'Ans Opt', 
                                                    'Explanation', cert_column_name]) # reading the sheet 1 which is named as Hyd Team Questions Addition and getting selected columns


    data.fillna('', inplace=True) # replacing all NaN with empty string 
    def _searchExamQuestion(row):
        global question_count
        try:
            text = row['Question']
            textQuery = r"{}".format(text)
            
            # print(f"question:{textQuery}")
            # print(type(textQuery))
            search_exam_question_url = f'{url}/support/content/search_exam_question?text_query={textQuery}' #url to search exam question
            search_response = requests.get(search_exam_question_url, **params)
            search_response_json = search_response.json()
            question_detail = search_response_json.get('questions')
            
            if question_detail: # if we are getting some response it means the question already exists then update it
                question_id = question_detail[0].get('question_id')
                question_assoc_data_url = f'{url}/support/content/question_assoc_data/{question_id}'
                question_assoc_data_response = requests.get(question_assoc_data_url, **params)
                question_assoc_data_response_json = question_assoc_data_response.json()
                already_mapped_cert = [i['cert']['short_name'] for i in question_assoc_data_response_json.get('cert_assocs')]
                question_count += 1
                
                if cert_name in already_mapped_cert:
                    print(f'{question_count}.Certificate name:{cert_name} is already mapped to the question with question_id:{question_id}')
                else:
                    _updateQuestionPlacement(question_id) 
            else: # if the question does not exists then adding it
                question_create_payload = {
                "question": {
                    "text": textQuery,
                    "answer": (row['Ans Opt'].split(' ')[-1]).lower(),
                    "explanation": row['Explanation'],
                    # "why_not": row['Why Not'],
                    "answer_options": {
                        'a':row["Option A"],
                        'b':row["Option B"],
                        'c':row["Option C"],
                        'd':row["Option D"]
                    }        
                },
                "cert_assocs": certificate_mapping(cert_name, str(row[cert_column_name])).get('cert_assocs')
            }
                # print(f'Logging payload', payload_json)


                payload_json = question_create_payload
                _addExamQuestion(payload_json)
        except Exception as e:
                print(f'Logging Exception while creating question:{e}')

    def _updateQuestionPlacement(question_id):
        global question_count
        try:
            update_question_url = f'{url}/support/content/assoc_question_cert'
            cert_assocs = certificate_mapping(cert_name, str(row[cert_column_name])).get('cert_assocs')[0]
            payload = {
                'question_id': question_id,
                'cert_id': cert_assocs.get('cert_id'),
                'cert_domain_id': cert_assocs.get('cert_domain_id'),
                'topic':cert_assocs.get('topic')
            }
            update_question_response = requests.post(update_question_url, json=payload, **params)
            update_question_response_json = update_question_response.json()
            question_count += 0
            print(f'{question_count}.Question with question id:{question_id} has been associated to the new cert:{update_question_response_json}')
        except Exception as e:
            print(f'Logging Exception while updating question:{e}')

    # 
    def _addExamQuestion(payload):
        global question_count
        try:
            add_exam_question_url = f'{url}/support/content/add_exam_question' # url to add exam question
            add_exam_question_response = requests.post(add_exam_question_url, json=payload, **params)
            add_exam_question_response_json = add_exam_question_response.json()
            question_count += 1  # 
            print(f"{question_count}.Question added successfully 🥳 with question id:{add_exam_question_response_json.get('question_id')}")
        except Exception as e:
            print(f'Logging Exception while creating question:{e}')
        
    for _, row in islice(data.iterrows(),None,None):
        if not row[cert_column_name]:
            continue
        else:
            _searchExamQuestion(row)
        
    sessionId = params.get('headers').get('ICSessionId')
    signout(sessionId)

else:
    print(f'Login not successful, try to change your credentials:{params}')