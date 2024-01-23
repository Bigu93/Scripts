import csv
import asyncio
import aiohttp
import os
import json
import argparse
from datetime import datetime


async def return_name(session, product_id):
    url = (
        f"https://xxx.pl/ajax/projector.php?action=get&product={product_id}&get=product"
    )
    async with session.get(url) as response:
        if response.headers.get("Content-Type") == "text/html; charset=utf-8":
            response_text = await response.text()

            json_start = response_text.find('{"')
            json_end = response_text.rfind("}") + 1
            json_content = response_text[json_start:json_end]
            json_data = json.loads(json_content)
            return json_data["product"]["name"]
        else:
            return "Unknown Product"


def check_date(date_to_check, limit_date):
    date1 = datetime.strptime(date_to_check.split(" ")[0], "%Y-%m-%d")
    date2 = datetime.strptime(limit_date, "%Y-%m-%d")
    return date1 > date2


def ensure_directory(directory_path):
    if not os.path.exists(directory_path):
        os.makedirs(directory_path)


def format_opinion_content(content):
    """
    Formats the given string to ensure it is enclosed in double quotes and internal double quotes are escaped.

    This function trims the input string, checks if it starts and ends with a double quote, and then processes it.
    If the string is already enclosed in double quotes, it escapes any internal double quotes and leaves the enclosing ones.
    If the string is not enclosed in double quotes, it adds them after escaping internal double quotes.

    Args:
    content (str): The string to be formatted.

    Returns:
    str: The formatted string, enclosed in double quotes, with internal double quotes escaped.
    """
    content = content.strip()

    if content.startswith('"') and content.endswith('"'):
        return '"' + content[1:-1].replace('"', '\\"') + '"'
    else:
        return '"' + content.replace('"', '\\"') + '"'


async def fetch_product_opinions():
    csv_folder = "csv"
    csv_file = "products_opinions.csv"
    date = "2023-01-01"
    ensure_directory(csv_folder)

    url_template = "https://xxx.pl/ajax/opinions.php?action=get&type=product&language=pol&resultsLimit=100&shopId=1&resultsPage={}&ordersBy[0][elementName]=date&ordersBy[0][sortDirection]=DESC"
    data = []

    async with aiohttp.ClientSession() as session:
        for i in range(10):
            url = url_template.format(i)
            async with session.get(url) as response:
                response_text = await response.text()

                json_start = response_text.find('{"')
                json_end = response_text.rfind("}") + 1
                json_content = response_text[json_start:json_end]

                print(f"URL: {url}")

                try:
                    json_data = json.loads(json_content)
                    opinions = json_data["results"]

                    for opinion in opinions:
                        if check_date(opinion["createDate"], date):
                            product_name = await return_name(
                                session, opinion["product"]["id"]
                            )
                            formatted_content = format_opinion_content(
                                opinion["content"]
                            )
                            data.append(
                                [
                                    opinion["orderSn"],
                                    opinion["createDate"],
                                    opinion["product"]["id"],
                                    product_name,
                                    opinion["rating"],
                                    formatted_content,
                                ]
                            )
                        else:
                            print(f"{opinion['createDate']} - before ", date)
                except json.JSONDecodeError:
                    print(f"Error decoding JSON from response for URL: {url}")

        with open(
            os.path.join(csv_folder, csv_file), "a", newline="", encoding="utf-8"
        ) as file:
            writer = csv.writer(file)
            for row in data:
                writer.writerow(row)


async def fetch_order_opinions():
    csv_folder = "csv"
    csv_file = "orders_opinions.csv"
    date = "2023-01-01"

    ensure_directory(csv_folder)

    url_template = "https://butosklep.pl/ajax/opinions.php?action=get&type=order&language=pol&resultsLimit=100&shopId=1&resultsPage={}&ordersBy[0][elementName]=date&ordersBy[0][sortDirection]=DESC"
    data = []

    async with aiohttp.ClientSession() as session:
        for i in range(10):
            url = url_template.format(i)
            async with session.get(url) as response:
                response_text = await response.text()

                json_start = response_text.find('{"')
                json_end = response_text.rfind("}") + 1
                json_content = response_text[json_start:json_end]

                print(f"URL: {url}")

                try:
                    json_data = json.loads(json_content)
                    opinions = json_data["results"]

                    for opinion in opinions:
                        if check_date(opinion["createDate"], date):
                            formatted_content = format_opinion_content(
                                opinion["content"]
                            )
                            data.append(
                                [
                                    opinion["orderSn"],
                                    opinion["createDate"],
                                    opinion["rating"],
                                    formatted_content,
                                ]
                            )
                        else:
                            print(f"{opinion['createDate']} - before ", date)
                except json.JSONDecodeError:
                    print(f"Error decoding JSON from response for URL: {url}")

        with open(
            os.path.join(csv_folder, csv_file), "a", newline="", encoding="utf-8"
        ) as file:
            writer = csv.writer(file)
            for row in data:
                writer.writerow(row)


parser = argparse.ArgumentParser(description="Fetch opinions from e-shop.")
parser.add_argument(
    "type",
    choices=["product", "order"],
    help="Type of opinions to fetch: product or order",
)
args = parser.parse_args()


if args.type == "product":
    asyncio.run(fetch_product_opinions())
elif args.type == "order":
    asyncio.run(fetch_order_opinions())
