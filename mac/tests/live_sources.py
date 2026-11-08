"""Probe public sources with the actual Mac helper. No Excel execution is implied."""
from pathlib import Path
import json
import subprocess

ROOT = Path(__file__).resolve().parents[2]
HELPER = ROOT / "mac/SMFMac.applescript"
SOURCES = [
    ("Finviz", "YAHOO", "https://finviz.com/quote.ashx?t=AAPL", ["Earnings", "Market Cap"]),
    ("Yahoo history", "GET", "https://query1.finance.yahoo.com/v8/finance/chart/AAPL?range=1mo&interval=1d", ["timestamp", "indicators"]),
    ("Yahoo options", "YAHOO", "https://query1.finance.yahoo.com/v7/finance/options/AAPL", ["expirationDates", "strikes"]),
    ("FRED", "GET", "https://fred.stlouisfed.org/data/DGS10", ["2026-"]),
    ("Zacks", "GET", "https://www.zacks.com/stock/quote/AAPL", ["AAPL"]),
]


def main():
    results = []
    for name, method, url, markers in SOURCES:
        try:
            reply = subprocess.run(["/usr/bin/osascript", str(HELPER), method + "\n" + url],
                                   capture_output=True, text=True, timeout=90, check=True)
            protocol, status, body = reply.stdout.removesuffix("\n").split("\n", 2)
            if protocol != "SMFHTTP/1":
                raise ValueError("Unexpected response format")
            result = {"source": name, "url": url, "http_status": int(status),
                      "body_characters": len(body),
                      "expected_markers_present": all(marker in body for marker in markers)}
            if status == "0":
                result["error"] = body[:300]
        except Exception as error:
            result = {"source": name, "url": url, "error": str(error)[:300]}
        results.append(result)
        print(json.dumps(result), flush=True)
    (ROOT / "mac/live-source-results.json").write_text(json.dumps(results, indent=2) + "\n")


if __name__ == "__main__":
    main()
