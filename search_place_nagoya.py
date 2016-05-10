# coding=utf-8

import time
import os
import configparser
import gspread
import urllib.request
import urllib.parse

from bs4 import BeautifulSoup
from oauth2client.service_account import ServiceAccountCredentials


def get_soup(URL):
    html = urllib.request.urlopen(URL)
    return BeautifulSoup(html, "html.parser")


def get_place_urls(BASEURL):
    soup = get_soup(BASEURL)
    table_rows = soup.find_all("span", class_="green s")
    place_urls = []
    for table_row in table_rows:
        reurl = table_row.parent.a.get("href")
        palace_url_params = urllib.parse.urlparse(reurl)

        if palace_url_params[4] != "":
            place_url = urllib.parse.urljoin(BASEURL, reurl)
            place_name = table_row.parent.a.string
            place_urls.append({"name": place_name, "url": place_url})
    return place_urls


def get_room_urls(place_urls):
    room_urls = []
    for place_url in place_urls:
        soup = get_soup(place_url["url"])
        tables = soup.find_all("table", class_="empty01")
        rooms = tables[0].find_all("th", class_="roomth")
        for room in rooms:
            room_url = urllib.parse.urljoin(place_url["url"], room.a.get("href"))
            room_name = room.a.string
            room_datas = room.parent.find_all("td")
            room_urls.append(
                {
                    "place_name": place_url["name"],
                    "place_url": place_url["url"],
                    "room_name": room_name,
                    "room_url": room_url,
                    "capacity": room_datas[0].string,
                    "am_fee": room_datas[1].string,
                    "pm_fee": room_datas[2].string,
                    "night_fee": room_datas[3].string,
                    "calender": {}
                }
            )
    return room_urls


def get_calender(room_urls):
    for room_url in room_urls:
        f_url = get_open_date(room_url["room_url"])
        for url in f_url:
            url_dict = urllib.parse.parse_qs(url)
            date = url_dict["year"][0] + "/" + url_dict["month"][0]
            soup = get_soup(url)
            tables = soup.find_all("table", class_="empty02")
            calender = [[], [], [], [], []]
            for table in tables:
                c = 0
                for row in table.find_all("tr"):
                    for col in row.find_all(["th", "td"])[1:]:
                        value = col.string
                        if col.name == "td" and c >= 2:
                            if col.img.attrs["src"] == "img/mark01.gif":
                                value = "○"
                            elif col.img.attrs["src"] == "img/mark02.gif":
                                value = "×"
                            elif col.img.attrs["src"] == "img/mark03.gif":
                                value = "△"
                            elif col.img.attrs["src"] == "img/mark04.gif":
                                value = "□"
                            elif col.img.attrs["src"] == "img/mark05.gif":
                                value = "◆"
                            else:
                                value = "-"
                        calender[c].append(value)
                    c += 1
            # calender.append(bgcolor)
            calender.append(url)
            room_url["calender"][date] = calender
    return room_urls


def get_open_date(room_url):
    urls = []
    soup = get_soup(room_url)
    datas = soup.find("div", class_="institution02").find("div", class_="datelink")
    for data in datas.find_all("li"):
        if not data.a:
            urls.append(room_url)
        else:
            urls.append(urllib.parse.urljoin(room_url, data.a.attrs["href"]))
    return urls


def get_gdoc():
    base = os.path.dirname(os.path.abspath(__file__))
    config_file_path = os.path.normpath(os.path.join(base, 'config.ini'))
    conf = configparser.ConfigParser()
    conf.read(config_file_path)
    json_api_key_file_path = os.path.normpath(os.path.join(base, conf.get("googledoc", "json_api_key_file_name")))

    scope = ['https://spreadsheets.google.com/feeds']
    credentials = ServiceAccountCredentials.from_json_keyfile_name(json_api_key_file_path, scope)
    gs = gspread.authorize(credentials)
    doc = gs.open_by_key(conf.get("googledoc", "doc_id"))
    return doc


def send_spreadsheet(calenders):
    doc = get_gdoc()
    lt = 3

    ws_list = doc.worksheets()
    for ws in ws_list:
        if ws._title != "README":
            doc.del_worksheet(ws)

    for k, v in calenders[0]["calender"].items():
        sheet = doc.add_worksheet(k, 200, 200)
        cell_list = sheet.range("A2:AK2")
        title = ["施設名", "部屋名", "定員", "午前の使用料", "午後の使用料", "夜間の使用料"]

        for day, weekday in zip(v[0], v[1]):
            date = k + "/" + day + "(" + weekday + ")"
            title.append(date)

        for (c, t) in zip(cell_list, title):
            c.value = t
        sheet.update_cells(cell_list)

        for l, calender in enumerate(calenders):
            line_data_list = [
                '=HYPERLINK("' + calender["place_url"] + '","' + calender['place_name']  + '")',
                calender['room_name'],
                calender['capacity'],
                calender['am_fee'],
                calender['pm_fee'],
                calender['night_fee']]

            for i, (am, pm, night) in enumerate(
                    zip(calender["calender"][k][2], calender["calender"][k][3], calender["calender"][k][4])):
                line_data_list.append(
                    '=HYPERLINK("' + calender["calender"][k][5] + '","' + am + ' : ' + pm + ' : ' + night + '")')

            cal_cell_range = "A" + str(l + lt) + ":AK" + str(l + lt)
            cal_cell_list = sheet.range(cal_cell_range)
            for (cell_data, line_data) in zip(cal_cell_list, line_data_list):
                cell_data.value = line_data

            sheet.update_cells(cal_cell_list)


if __name__ == '__main__':
    #start = time.time()
    BASEURL = "https://www.suisin.city.nagoya.jp/system/institution/index.cgi"
    place_urls = get_place_urls(BASEURL)
    # del place_urls[1:-1]
    room_urls = get_room_urls(place_urls)
    # del room_urls[1:-1]
    calender = get_calender(room_urls)
    send_spreadsheet(calender)
    #elapsed_time = time.time() - start
    # print("elapsed_time:{0}".format(elapsed_time) + "[sec]")
