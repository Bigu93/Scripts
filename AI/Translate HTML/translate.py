from dotenv import load_dotenv
from openai import OpenAI
from bs4 import BeautifulSoup, NavigableString, Tag
import os
import re

load_dotenv()
client = OpenAI(api_key=os.getenv("OPEN_API_KEY"))


def translate_text(
    text,
    target_language,
    source_language="Polski",
):
    if not text.strip() or not re.search("[\w]", text):
        return text
    try:
        response = client.chat.completions.create(
            model="gpt-3.5-turbo-1106",
            messages=[
                {
                    "role": "system",
                    "content": f"Zadanie: Przetłumacz wiadomość użytkownika z języka {source_language}iego na {target_language}. Nie tłumacz danych osobowych i adresowych. Zwróć tylko tłumaczenie, nic więcej.",
                },
                {"role": "user", "content": text},
            ],
            max_tokens=4000,
            temperature=0.5,
        )
        print(
            f"{source_language}: {text.strip()} -{target_language}: {response.choices[0].message.content.strip()}\n"
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        print(f"An error occurred: {e}")
        return text


def translate_html_section(section):
    if isinstance(section, Tag):
        if section.get("data-translated") == "true":
            # Skip already translated sections
            return

        for child in section.contents:
            translate_html_section(child)

        section["data-translated"] = "true"  # Mark as translated
    elif isinstance(section, NavigableString):
        translated_text = translate_text(str(section))
        section.replace_with(translated_text)


def split_and_translate_html(html_content):
    soup = BeautifulSoup(html_content, "html.parser")
    sections = soup.find_all(
        ["div", "h1", "h2", "h3", "h4", "h5", "h6", "p", "ol", "li"]
    )  # Add more tags as needed

    for section in sections:
        translate_html_section(section)

    return str(soup)


def remove_data_translated_attribute(soup):
    for tag in soup.find_all(True, {"data-translated": "true"}):
        del tag["data-translated"]


def save_html_to_file(html_content, language):
    file_name = f"translated_{language}.html"
    try:
        with open(file_name, "w", encoding="utf-8") as file:
            file.write(html_content)
        print(f"File saved as {file_name}")
    except Exception as e:
        print(f"An error occurred while saving the file: {e}")


html_content = """
TEXT TO TRANSLATE
"""

translated_html = split_and_translate_html(html_content)
soup = BeautifulSoup(translated_html, "html.parser")
remove_data_translated_attribute(soup)
final_html = str(soup)
save_html_to_file(final_html, "OUTPUT_LANG")
