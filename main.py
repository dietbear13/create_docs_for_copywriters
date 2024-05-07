import json
import os
import xml.etree.ElementTree as ET
from datetime import datetime

import pandas as pd
import requests
from bs4 import BeautifulSoup
from google.oauth2 import service_account
from googleapiclient.discovery import build
from openpyxl import load_workbook


def load_config(config_file_path):
    with open(config_file_path, 'r') as config_file:
        config = json.load(config_file)
    return config


config = load_config('config.json')
folder_id = config['folder_id']
service_account_key_path = config['service_account_key_path']
xmlriver_url = config['xmlriver_url']

SCOPES = ['https://www.googleapis.com/auth/documents']
creds = None
if os.path.exists(service_account_key_path):
    creds = service_account.Credentials.from_service_account_file(service_account_key_path)
    service = build('docs', 'v1', credentials=creds)
    drive_service = build('drive', 'v3', credentials=creds)

service = build('docs', 'v1', credentials=creds)


def extract_h2_titles(html_content):
    soup = BeautifulSoup(html_content, 'html.parser')
    h2_titles = [h2.text for h2 in soup.find_all('h2')]
    h2_titles = [title.replace('\xa0', ' ').replace('\r', '')
                 .replace('\n', '').replace('\t', '') for title in h2_titles]
    return h2_titles


def extract_meta_titles(html_content):
    soup = BeautifulSoup(html_content, 'html.parser')
    meta_titles = [title.text for title in soup.find_all('title')]
    meta_titles = [title.replace('\xa0', ' ').replace('\r', '')
                   .replace('\n', '').replace('\t', '') for title in meta_titles]
    return meta_titles


def create_document_and_write_to_file(topic, output_my_headers, parse_meta_titles, parse_h2, links):
    body = {
        'title': topic.capitalize()
    }
    doc = service.documents().create(body=body).execute()
    document_id = doc.get('documentId')
    print(f'Создан новый документ: {doc.get("title")}')

    title_section = f"Title: {topic.capitalize()}\n\nTitle в выдаче:\n{"\n".join(parse_meta_titles)}\n\n"
    content_section = f"Содержание текста\n{output_my_headers}\n\n"
    competitors_section = f"У конкурентов:\n{"\n".join(parse_h2)}\n"
    technical_requirements_section = "Технические требования\nОбъем текста от X слов. Объем может быть больше или меньше, главное – раскрыть тему полностью, но без воды.\nУникальность по text.ru от 85%.\nИзбегаем речевого мусора, канцеляритов и вводных слов.\n"
    examples_section = f"Примеры текстов:\n"

    pattern = title_section + content_section + competitors_section + technical_requirements_section + examples_section

    requests = [
        {
            'insertText': {
                'location': {
                    'index': 1
                },
                'text': f"{pattern}"
            }
        }
    ]

    start_index = pattern.index("Title в выдаче:") + len("Title в выдаче: ")
    end_index = pattern.index("Title в выдаче:") + len("Title в выдаче: ") + len(parse_meta_titles)

    requests.append({
        'createParagraphBullets': {
            'range': {
                'startIndex': start_index + 1,
                'endIndex': end_index
            },
            'bulletPreset': 'BULLET_DISC_CIRCLE_SQUARE'
        }
    })

    pattern_length = len(pattern)
    list_items = links

    start_index = pattern.index("Title в выдаче:") + len("Title в выдаче: ")
    end_index = pattern.index("Содержание текста")
    requests.append({
        'createParagraphBullets': {
            'range': {
                'startIndex': start_index + 1,
                'endIndex': end_index + 1
            },
            'bulletPreset': 'BULLET_DISC_CIRCLE_SQUARE'
        }
    })

    start_index = pattern.index("У конкурентов:") + len("У конкурентов: ")
    end_index = pattern.index("Технические требования")
    requests.append({
        'createParagraphBullets': {
            'range': {
                'startIndex': start_index + 1,
                'endIndex': end_index + 1
            },
            'bulletPreset': 'BULLET_DISC_CIRCLE_SQUARE'
        }
    })

    start_index = pattern.index("Технические требования") + len("Технические требования ")
    end_index = pattern.index("Примеры текстов")
    requests.append({
        'createParagraphBullets': {
            'range': {
                'startIndex': start_index + 1,
                'endIndex': end_index + 1
            },
            'bulletPreset': 'BULLET_DISC_CIRCLE_SQUARE'
        }
    })

    for item in list_items:
        requests.append({
            'insertText': {
                'location': {
                    'index': pattern_length + 1
                },
                'text': f"{item}\n",
            }
        })
        pattern_length += len(item) + 1

    headings = ["Содержание текста", "У конкурентов:", "Технические требования", "Примеры текстов"]
    for heading in headings:
        start_index = pattern.index(heading) + 1
        end_index = start_index + len(heading)
        requests.append({
            'updateParagraphStyle': {
                'range': {
                    'startIndex': start_index,
                    'endIndex': end_index
                },
                'paragraphStyle': {
                    'namedStyleType': 'HEADING_2',
                },
                'fields': 'namedStyleType'
            }
        })

    start_index = pattern.index("Объем текста от X слов.") + len("Объем текста от ") + 1
    end_index = start_index + 1
    requests.append({
        'updateTextStyle': {
            'range': {
                'startIndex': start_index,
                'endIndex': end_index,
            },
            'textStyle': {
                'backgroundColor': {
                    'color': {
                        'rgbColor': {'red': 0.98, 'green': 0.97, 'blue': 0.55}
                    }
                }
            },
            'fields': 'backgroundColor'
        }
    })

    start_index = pattern.index("Title:") + len("Title: ") + 1
    end_index = start_index + len(topic)
    requests.append({
        'updateTextStyle': {
            'range': {
                'startIndex': start_index,
                'endIndex': end_index,
            },
            'textStyle': {
                'backgroundColor': {
                    'color': {
                        'rgbColor': {'red': 0.98, 'green': 0.97, 'blue': 0.55}
                    }
                }
            },
            'fields': 'backgroundColor'
        }
    })

    service.documents().batchUpdate(documentId=document_id, body={'requests': requests}).execute()
    file_id = doc['documentId']
    print(f'Текст успешно записан в документ: {topic}')

    file = drive_service.files().get(fileId=file_id, fields='parents').execute()
    previous_parents = ",".join(file.get('parents'))
    file = drive_service.files().update(fileId=document_id, addParents=folder_id, removeParents=previous_parents,
                                        fields='id, parents').execute()
    link = f"https://docs.google.com/document/d/{file.get('id')}"

    try:
        df = pd.read_excel('result_links.xlsx')
    except FileNotFoundError:
        df = pd.DataFrame(columns=['Тема', 'Ссылка'])

    current_date = datetime.now().strftime('%d.%m.%Y')

    new_row = {'Тема': topic, 'Ссылка': link, 'Дата': current_date}
    df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)

    df.to_excel('result_links.xlsx', index=False)

    print(f"\nСтатья «{topic}» — {link}")
    return topic, link


parse_h2 = []
links = []
parse_meta_titles = []
my_h2 = []


def parse_google_results(query):
    url = f'{xmlriver_url}&query={query}'

    try:
        response = requests.get(url)

        if response.status_code == 200:
            root = ET.fromstring(response.content)
            count = 0
            for group in root.findall(".//group"):
                url = group.find(".//url").text
                links.append(url)
                count += 1
                if count >= 5:
                    break

            h2_titles = []

            for link in links:
                try:
                    response = requests.get(link)
                    if response.status_code == 200:
                        h2_titles.extend(extract_h2_titles(response.content))
                        parse_meta_titles.extend(extract_meta_titles(response.content))
                except requests.exceptions.RequestException as e:
                    print(f"Не удалось получить контент страницы: {link}")
                    continue

            print(parse_meta_titles)
            print(h2_titles)

            return h2_titles, parse_meta_titles, links
    except Exception as e:
        print(f"Произошла ошибка при выполнении запроса: {e}")
        return [], [], []


def interact():
    output_my_headers = ""

    wb = load_workbook('input_table.xlsx')
    sheet = wb.active
    for cell in sheet['A']:
        topic = cell.value
        if topic:
            my_h1 = topic.split(': ')[0]
            my_h2_titles = topic.split(': ')[1]
            my_h2_splited = my_h2_titles.split(', ')

            output_my_headers += f"H1 содержит «{my_h1.capitalize()}»\n"
            for my_h2_title in my_h2_splited:
                output_my_headers += f"H2: {my_h2_title.capitalize()}\n"

            parse_h2_titles, _, links = parse_google_results(topic.split(':')[0])
            for h2_title in parse_h2_titles:
                parse_h2.append(h2_title)

            create_document_and_write_to_file(my_h1, output_my_headers, parse_meta_titles, parse_h2, links)
            output_my_headers = ""
            parse_h2.clear()
            links.clear()
            parse_meta_titles.clear()

    print("Генерация завершена.")


interact()
