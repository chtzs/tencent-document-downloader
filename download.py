import requests
import re
import json
import datetime
import load_cookies

OPENDOC_API = "https://docs.qq.com/dop-api/opendoc"


class UserData:
    def __init__(self):
        self.cookies = None

    def set_cookies(self, cookiesfile):
        self.cookies = load_cookies.load_cookies(cookiesfile)

    def get_cookies(self):
        cookie_strings = []
        if self.cookies != None:
            for cookie in list(self.cookies):
                cookie_strings.append(cookie.name + '=' + cookie.value)
        cookie_headers = {'cookie': '; '.join(cookie_strings)}
        # cookie_headers = {'cookie': cookie_string}
        return cookie_headers


class SheetDownloader:
    def __init__(self, url: str, cookie_data: UserData = UserData()) -> None:
        self.url = url
        self.cookie_data = cookie_data
        self.title = ""
        self.tabs = []
        self._init_params()
        self._fetch_doc_detail()

    def _init_params(self) -> None:
        t = datetime.datetime.timestamp(datetime.datetime.now())
        # In million secs
        t = int(t * 1000)
        s = self.url.split("?")
        referer = s[0]
        id = referer.split("sheet/")[1]
        self.opendoc_params = {
            "id": id,
            "noEscape": "1",
            "normal": "1",
            "outformat": "1",
            "startrow": "0",
            "endrow": "60",
            "wb": "1",
            "nowb": "0",
            "callback": "clientVarsCallback",
            "xsrf": "",
            "t": t
        }
        self.header = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/103.0.0.0 Safari/537.36",
            "Referer": referer
        }
        self.header.update(self.cookie_data.get_cookies())

    def _fetch_doc_detail(self):
        opendoc_json = self._fetch_doc_json()
        self.title = opendoc_json["clientVars"]["title"]
        self.tabs = opendoc_json["clientVars"]["collab_client_vars"]["header"][0]["d"]

    def _fetch_doc_json(self, params={}):
        cloned_params = self.opendoc_params.copy()
        cloned_params.update(params)
        opendoc_text = requests.get(
            OPENDOC_API, headers=self.header, params=cloned_params).text
        json_content = opendoc_text[len("clientVarsCallback("):-1]
        opendoc_json = json.loads(json_content)

        return opendoc_json

    def fetch_sheet_data(self, tab_id: str):
        opendoc_json = self._fetch_doc_json(params={
            "tab": tab_id
        })
        max_row = opendoc_json["clientVars"]["collab_client_vars"]["maxRow"]
        max_col = opendoc_json["clientVars"]["collab_client_vars"]["maxCol"]
        return opendoc_json["clientVars"]["collab_client_vars"]["initialAttributedText"]["text"][0], max_row, max_col
