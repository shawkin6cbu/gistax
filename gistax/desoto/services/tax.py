import requests, re
from bs4 import BeautifulSoup

WELCOME = "http://www.desotoms.info/Webpgms/welcome.pgm"
CGI     = "http://www.desotoms.info/Webpgms/aptaxest7.pgm"
UA_HDR  = {"User-Agent": "Mozilla/5.0"}

DISTRICT_MAP = {
    "County": 0,
    "Hernando": 1,
    "Horn Lake": 2,
    "Olive Branch": 3,
    "Southaven": 4,
    "Walls": 5,
}
DISTRICT_OPTIONS = list(DISTRICT_MAP.keys())

_MONEY_RE = re.compile(r"\$?([\d,]+\.\d{2})", re.A)

def fetch_total(value: str, district: str) -> str | None:
    with requests.Session() as s:
        s.get(WELCOME, headers=UA_HDR, timeout=10)
        payload = {"apprval": value,
                   "millage": DISTRICT_MAP[district],
                   "Calc": "Calculate"}
        r = s.post(CGI, headers=UA_HDR, data=payload, timeout=10)
        r.raise_for_status()
        soup = BeautifulSoup(r.text, "html.parser")
        img = soup.find("img", alt=re.compile("normal primary residence", re.I))
        if not img:
            return None
        row = img.find_parent("tr")
        if not row:
            return None
        cells = row.find_all("td")
        if not cells:
            return None
        m = _MONEY_RE.search(cells[-1].get_text())
        return m.group(1) if m else None
