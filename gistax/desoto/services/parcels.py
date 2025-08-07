import requests

PARCEL_URL = (
    "https://gis.desotocountyms.gov/arcgis/rest/services/"
    "CountyWebMap/County_Web_Map/MapServer/29/query"
)

FIELDS = (
    "FULL_ADDR,PIN,OWNER_NAME,SECOND_OWNER,"
    "CITY,STATE,ZIP_CODE,SUBD_NAME,LOT"
)

def query(prefix: str, limit: int = 10):
    prefix_sql = prefix.upper().replace("'", "''")
    where = f"UPPER(FULL_ADDR) LIKE '{prefix_sql}%'"
    try:
        r = requests.get(
            PARCEL_URL,
            params={
                "where": where,
                "outFields": FIELDS,
                "returnGeometry": "false",
                "returnDistinctValues": "true",
                "orderByFields": "FULL_ADDR",
                "resultRecordCount": limit,
                "f": "json",
            },
            timeout=15,
        )
        r.raise_for_status()
        return [f["attributes"] for f in r.json().get("features", [])]
    except Exception as exc:
        print("Parcel lookup failed:", exc)
        return []
