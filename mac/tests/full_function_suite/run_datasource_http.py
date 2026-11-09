"""Record bounded curl evidence for the exact inventoried SMF destinations."""
from __future__ import annotations

from concurrent.futures import ThreadPoolExecutor, as_completed
import argparse
from datetime import datetime, timezone
import hashlib
import json
from pathlib import Path
import re
import subprocess
import tempfile
import threading
from urllib.parse import urlsplit, urlunsplit, parse_qsl, urlencode, quote

ROOT = Path(__file__).resolve().parents[3]
INVENTORY = Path(__file__).with_name("datasource_cases.json")
META = Path(__file__).with_name("datasource_inventory_meta.json")
SUPPORT_REGISTRY = ROOT / "mac/not_supported.json"
OUTPUT: Path | None = None
LOCKS: dict[str, threading.Lock] = {}
GUARD = threading.Lock()


def load_support_registry(path: Path) -> dict:
    payload = path.read_bytes()
    source = json.loads(payload)
    if not isinstance(source, dict) or source.get("schema_version") != 1:
        raise ValueError("Unsupported support-registry schema")
    status = source.get("status")
    if not isinstance(status, str) or not status:
        raise ValueError("Support registry has no display status")
    entries = {}
    for entry in source.get("entries", []):
        if not isinstance(entry, dict) or not isinstance(entry.get("id"), str) or not entry["id"]:
            raise ValueError("Support registry has an invalid entry ID")
        if entry["id"] in entries or not isinstance(entry.get("reason"), str) or not entry["reason"]:
            raise ValueError("Support registry has a duplicate ID or missing reason")
        entries[entry["id"]] = entry
    rules = []
    for rule in source.get("url_rules", []):
        if not isinstance(rule, dict) or rule.get("id") not in entries:
            raise ValueError("Support URL rule has an unknown entry ID")
        if not isinstance(rule.get("host"), str) or not rule["host"] or not isinstance(rule.get("path"), str) or not rule["path"].startswith("/"):
            raise ValueError("Support URL rule needs a host and absolute path")
        rules.append((rule["id"], rule["host"].lower(), rule["path"].lower()))
    return {"status": status, "entries": entries, "url_rules": tuple(rules),
            "sha256": hashlib.sha256(payload).hexdigest(), "schema_version": source["schema_version"]}


def unsupported_url(url: str, support: dict) -> tuple[str, str] | None:
    """Match the VBA host/path policy; do not normalize unpublished path aliases."""
    scheme_at = url.find("://")
    if scheme_at < 0 or url[:scheme_at].lower() not in {"http", "https"}:
        return None
    authority_at = scheme_at + 3
    end_at = len(url)
    for pos in range(authority_at, len(url)):
        if url[pos] in "/?#":
            end_at = pos
            break
    authority = url[authority_at:end_at].rsplit("@", 1)[-1]
    if authority.startswith("["):
        return None
    authority = authority.split(":", 1)[0]
    if authority.endswith("."):
        authority = authority[:-1]
    if not authority:
        return None
    path = "/"
    if end_at < len(url) and url[end_at] == "/":
        path_at = end_at
        end_at = len(url)
        for pos in range(path_at, len(url)):
            if url[pos] in "?#":
                end_at = pos
                break
        path = url[path_at:end_at]
    for entry_id, host, route_path in support["url_rules"]:
        if authority.lower() == host and path.lower() == route_path:
            return entry_id, support["entries"][entry_id]["reason"]
    return None


def host_lock(url: str):
    host = urlsplit(url).hostname or "invalid"
    with GUARD:
        return LOCKS.setdefault(host, threading.Lock())


def curl(url: str, prefix: Path, cookie: Path | None = None) -> dict:
    body = prefix.with_suffix(".body")
    header = prefix.with_suffix(".headers")
    args = ["/usr/bin/curl", "--http1.1", "--silent", "--show-error", "--location",
            "--compressed", "--connect-timeout", "8", "--max-time", "25",
            "--proto", "=http,https", "--proto-redir", "=http,https",
            "--user-agent", "Mozilla/5.0 SMF-Mac/0.1", "--request", "GET",
            "--output", str(body), "--dump-header", str(header),
            "--write-out", "%{url_effective}\n%{http_code}\n%{content_type}\n%{size_download}\n%{num_redirects}\n%{time_total}\n"]
    if cookie:
        args += ["--cookie", str(cookie), "--cookie-jar", str(cookie)]
    args += ["--url", url]
    with host_lock(url):
        try:
            result = subprocess.run(args, capture_output=True, text=True, timeout=30)
            code, stdout, error = result.returncode, result.stdout, result.stderr.strip()
        except subprocess.TimeoutExpired as exc:
            code, stdout, error = 124, "", "Outer subprocess timeout: " + str(exc)
    parts = stdout.splitlines()
    payload = body.read_bytes() if body.exists() else b""
    text = payload.decode("utf-8", errors="replace")
    title = re.search(r"<title[^>]*>(.*?)</title>", text, re.I | re.S)
    blocking = [token for token in ("Access Denied", "access denied", "captcha", "Just a moment...",
                                   "cf-chl-", "Unauthorized", "Invalid Crumb", "Too Many Requests",
                                   "no longer available", "discontinued") if token.lower() in text.lower()]
    raw_headers = header.read_text(errors="replace") if header.exists() else ""
    # Keep header names and status evidence without retaining anonymous cookie values.
    if header.exists():
        header.write_text(re.sub(r"(?im)^(set-cookie:)[^\r\n]*", r"\1 [redacted]", raw_headers))
    return {"requested_url": url, "effective_url": parts[0] if parts else "",
            "http_status": int(parts[1]) if len(parts) > 1 and parts[1].isdigit() else 0,
            "content_type": parts[2] if len(parts) > 2 else "", "bytes": len(payload),
            "redirects": parts[4] if len(parts) > 4 else "", "curl_seconds": parts[5] if len(parts) > 5 else "",
            "curl_exit": code, "error": error, "body_sha256": hashlib.sha256(payload).hexdigest(),
            "body_path": str(body.relative_to(OUTPUT)), "headers_path": str(header.relative_to(OUTPUT)),
            "title": re.sub(r"\s+", " ", title.group(1)).strip()[:300] if title else "",
            "blocking_markers": blocking, "set_cookie_present": bool(re.search(r"(?im)^set-cookie:", raw_headers))}


def run_job(index: int, job: dict, support: dict | None = None) -> dict:
    prefix = OUTPUT / f"{index:03d}"
    url, mode = job["url"], job.get("mode", "GET")
    record = {"index": index, "source_ids": job["source_ids"], "mode": mode,
              "kind": job.get("kind", "runtime_data"), "started_utc": datetime.now(timezone.utc).isoformat()}
    support = load_support_registry(SUPPORT_REGISTRY) if support is None else support
    skipped = unsupported_url(url, support)
    if skipped is not None:
        support_id, reason = skipped
        record.update({"support_id": support_id, "support_status": support["status"],
                       "request_performed": False, "network_requests": 0})
        record["response"] = {"requested_url": url, "effective_url": "", "http_status": 0,
                              "support_id": support_id, "support_status": support["status"],
                              "request_performed": False, "network_requests": 0,
                              "curl_exit": None, "error": support["status"] + " -- " + reason,
                              "body_path": "", "headers_path": "", "body_sha256": "", "bytes": 0,
                              "content_type": "", "title": "", "blocking_markers": [],
                              "set_cookie_present": False}
    elif mode == "YAHOO" and url.startswith(("https://query1.finance.yahoo.com/", "https://query2.finance.yahoo.com/")):
        with tempfile.TemporaryDirectory(prefix="smf-probe-cookies-") as temp:
            cookie = Path(temp) / "cookies.txt"
            first = curl("https://fc.yahoo.com", OUTPUT / f"{index:03d}-cookie", cookie)
            crumb = curl("https://query1.finance.yahoo.com/v1/test/getcrumb", OUTPUT / f"{index:03d}-crumb", cookie)
            crumb_value = (OUTPUT / crumb["body_path"]).read_text(errors="replace") if (OUTPUT / crumb["body_path"]).exists() else ""
            record["auth"] = {"cookie_status": first["http_status"], "cookie_received": first["set_cookie_present"],
                              "crumb_status": crumb["http_status"], "crumb_valid": False}
            valid = crumb["http_status"] == 200 and 0 < len(crumb_value) <= 100 and not any(x in crumb_value for x in "\r\n{<")
            record["auth"]["crumb_valid"] = valid
            # Crumb bodies are anonymous session material; their values are not needed in the report.
            (OUTPUT / crumb["body_path"]).write_text("[crumb redacted]\n")
            if valid:
                parsed = urlsplit(url)
                params = [(k, v) for k, v in parse_qsl(parsed.query, keep_blank_values=True) if k != "crumb"]
                # Preserve the source URL's query bytes; only replace/add its crumb parameter.
                existing = "&".join(part for part in parsed.query.split("&") if part and not part.startswith("crumb="))
                query = existing + ("&" if existing else "") + "crumb=" + quote(crumb_value, safe="-._~")
                actual_url = urlunsplit((parsed.scheme, parsed.netloc, parsed.path, query, parsed.fragment))
                record["response"] = curl(actual_url, prefix, cookie)
                record["response"]["requested_url"] = url
                record["response"]["effective_url"] = re.sub(r"([?&]crumb=)[^&]*", r"\1[redacted]", record["response"]["effective_url"])
            else:
                record["response"] = {"requested_url": url, "http_status": 0, "curl_exit": None,
                                      "error": "SMF Yahoo handshake failed before the data request", "body_path": ""}
    else:
        record["response"] = curl(url, prefix)
    record["finished_utc"] = datetime.now(timezone.utc).isoformat()
    (OUTPUT / f"{index:03d}.json").write_text(json.dumps(record, indent=2) + "\n")
    return record


def main():
    global OUTPUT, INVENTORY, META
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, required=True, help="New folder for this HTTP run; paths in records are relative to this folder.")
    parser.add_argument("--workers", type=int, default=3)
    parser.add_argument("--inventory", type=Path, default=INVENTORY)
    parser.add_argument("--metadata", type=Path, default=META)
    parser.add_argument("--support-registry", type=Path, default=SUPPORT_REGISTRY,
                        help="Reviewed exact retired-route registry; read before any curl or Yahoo handshake.")
    args = parser.parse_args()
    if not 1 <= args.workers <= 6:
        parser.error("--workers must be between 1 and 6")
    INVENTORY = args.inventory.resolve()
    META = args.metadata.resolve()
    OUTPUT = args.output.resolve()
    support_path = args.support_registry.resolve()
    support = load_support_registry(support_path)
    OUTPUT.mkdir(parents=True, exist_ok=False)
    inventory = json.loads(INVENTORY.read_text())
    meta = json.loads(META.read_text())
    jobs: dict[tuple[str, str], dict] = {}
    def add(url, source, kind, mode="GET"):
        key = (url, mode)
        if key not in jobs:
            jobs[key] = {"url": url, "mode": mode, "kind": kind, "source_ids": []}
        if source not in jobs[key]["source_ids"]:
            jobs[key]["source_ids"].append(source)
    for provider in inventory:
        for url in provider["probe_urls"]:
            add(url, provider["id"], provider["kind"], provider.get("mode", "GET"))
        for extra in provider.get("extra_probes", []):
            # Crumb is exercised with the required cookie session in each Yahoo mode job.
            if extra["id"] != "yahoo_crumb":
                add(extra["url"], extra["id"], extra["kind"], extra.get("mode", "GET"))
    for ref in meta.get("unreachable_option_references", []):
        for url in ref.get("probe_urls", []):
            add(url, "unreachable:" + ref["function"], "unreachable_legacy_reference")
    records = []
    (OUTPUT / "run-metadata.json").write_text(json.dumps({"schema_version": 2, "started_utc": datetime.now(timezone.utc).isoformat(), "path_base": "this_output_folder", "inventory_sha256": hashlib.sha256(INVENTORY.read_bytes()).hexdigest(), "support_registry_sha256": support["sha256"], "support_registry_schema_version": support["schema_version"], "workers": args.workers}, indent=2) + "\n")
    print(json.dumps({"output": str(OUTPUT), "jobs": len(jobs), "workers": args.workers}), flush=True)
    with ThreadPoolExecutor(max_workers=args.workers) as pool:
        futures = [pool.submit(run_job, index, job, support) for index, job in enumerate(jobs.values(), 1)]
        for future in as_completed(futures):
            record = future.result(); records.append(record)
            r = record["response"]
            print(json.dumps({"index": record["index"], "sources": record["source_ids"],
                              "status": r["http_status"], "exit": r.get("curl_exit"), "bytes": r.get("bytes", 0),
                              "support_id": r.get("support_id", ""), "request_performed": r.get("request_performed")}), flush=True)
    records.sort(key=lambda item: item["index"])
    (OUTPUT / "results.json").write_text(json.dumps(records, indent=2) + "\n")
    print(json.dumps({"complete": True, "jobs": len(records),
                      "skipped_not_supported": sum(record.get("request_performed") is False for record in records),
                      "results": str(OUTPUT / "results.json")}), flush=True)


if __name__ == "__main__":
    main()
