#!/usr/bin/env python3
"""
Keep the trip site's changing numbers fresh: exchange rates, weather, ticket prices.

Run from the project folder:

  python3 tools/update_trip.py                 # check everything, change nothing
  python3 tools/update_trip.py --apply         # also write new rates (and forecast temps) into index.html
  python3 tools/update_trip.py --set "ESCAPE day pass=180"         # change one price
  python3 tools/update_trip.py --set "Flight Penang → KL=95-240"   # a price range

Then commit and push as usual. Needs only Python 3 and an internet connection.
Free sources, no keys: open.er-api.com (rates) and open-meteo.com (weather forecast,
which only reaches about 16 days ahead, so it fills in as the trip gets closer).
"""
import argparse, datetime as dt, json, re, sys, urllib.parse, urllib.request
from pathlib import Path

SITE = Path(__file__).resolve().parent.parent / "index.html"
CITIES = [  # name on the weather card, coordinates, first and last day there
    ("Langkawi", 6.35, 99.80, "2026-11-27", "2026-12-01"),
    ("Penang", 5.41, 100.33, "2026-12-01", "2026-12-05"),
    ("Kuala Lumpur", 3.15, 101.71, "2026-12-05", "2026-12-10"),
]


def get_json(url):
    req = urllib.request.Request(url, headers={"User-Agent": "trip-site-updater"})
    with urllib.request.urlopen(req, timeout=20) as r:
        return json.load(r)


def rates(html, apply):
    print("\n== Exchange rates (RM1 =) ==")
    m = re.search(r"var RATES = \{MYR:1, EGP:([\d.]+), GBP:([\d.]+), USD:([^}]+)\};", html)
    if not m:
        print("  Couldn't find RATES in index.html"); return html
    old = {"EGP": float(m.group(1)), "GBP": float(m.group(2)), "USD": float(eval(m.group(3), {}))}
    try:
        d = get_json("https://open.er-api.com/v6/latest/MYR")["rates"]
    except Exception as e:
        print("  Couldn't fetch rates:", e); return html
    new = {k: d[k] for k in ("EGP", "GBP", "USD")}
    for k in new:
        ch = (new[k] - old[k]) / old[k] * 100
        print(f"  {k}: {old[k]:.4f} -> {new[k]:.4f}  ({ch:+.1f}%)")
    if apply:
        html = html.replace(m.group(0), "var RATES = {MYR:1, EGP:%s, GBP:%s, USD:%s};" % (
            round(new["EGP"], 4), round(new["GBP"], 6), round(new["USD"], 6)))
        print("  Written to index.html")
    return html


def weather(html, apply):
    print("\n== Weather forecast for your dates ==")
    today = dt.date.today()
    for name, lat, lon, a, b in CITIES:
        start, end = dt.date.fromisoformat(a), dt.date.fromisoformat(b)
        last = min(end, today + dt.timedelta(days=14))
        if last < max(start, today):
            print(f"  {name}: not in forecast range yet (opens around {start - dt.timedelta(days=14)})"); continue
        q = urllib.parse.urlencode({"latitude": lat, "longitude": lon, "timezone": "Asia/Kuala_Lumpur",
            "daily": "temperature_2m_max,temperature_2m_min,precipitation_probability_max",
            "start_date": max(start, today).isoformat(), "end_date": last.isoformat()})
        try:
            d = get_json("https://api.open-meteo.com/v1/forecast?" + q)["daily"]
        except Exception as e:
            print(f"  {name}: couldn't fetch ({e})"); continue
        lo, hi = round(min(d["temperature_2m_min"])), round(max(d["temperature_2m_max"]))
        print(f"  {name}: {lo}–{hi}°C")
        for day, mn, mx, p in zip(d["time"], d["temperature_2m_min"], d["temperature_2m_max"], d["precipitation_probability_max"]):
            print(f"     {day}  {round(mn)}–{round(mx)}°C  rain chance {p}%")
        if apply:
            pat = re.compile(r'(<h3>' + re.escape(name) + r' · [^<]*</h3>.*?<p class="temp">)[^<]*(</p>)', re.S)
            html, n = pat.subn(lambda m: f"{m.group(1)}{lo}–{hi}°C{m.group(2)}", html, count=1)
            print("     card updated" if n else "     (weather card not found)")
    return html


def prices(html, sets):
    print("\n== Prices in the day plans (RM per person) ==")
    items = re.findall(r'\["([^"\]]+)",(\d+)(?:,(\d))?(?:,(\d+))?\]', html)
    seen = []
    for name, lo, _, hi in items:
        if name in seen: continue
        seen.append(name)
        val = lo + ("–" + hi if hi else "")
        link = "https://www.google.com/search?q=" + urllib.parse.quote(name + " price Malaysia foreigner")
        print(f"  {name:40s} RM{val:9s} {link}")
    for s in sets or []:
        name, _, val = s.partition("=")
        name, val = name.strip(), val.strip()
        m = re.fullmatch(r"(\d+)(?:-(\d+))?", val)
        if not name or not m:
            print(f"  Skipped '{s}': use \"Name=120\" or \"Name=90-260\""); continue
        pat = re.compile(r'\["' + re.escape(name) + r'",\d+(?:,(\d))?(?:,\d+)?\]')
        rep = lambda mm: '["%s",%s%s]' % (name, m.group(1), (",%s,%s" % (mm.group(1) or 0, m.group(2))) if m.group(2) else ("," + mm.group(1) if mm.group(1) else ""))
        html, n = pat.subn(rep, html)
        print(f"  {name}: set to RM{val} ({n} place{'s' if n != 1 else ''})" if n else f"  '{name}' not found; copy the name exactly from the list above")
    if sets:
        month = dt.date.today().strftime("%b %Y")
        html = re.sub(r'(<span id="pricesChecked">)[^<]*(</span>)', r"\g<1>" + month + r"\g<2>", html)
    return html


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--apply", action="store_true", help="write new rates and forecast temperatures")
    ap.add_argument("--set", action="append", metavar='"Name=price"', help="change a price (repeatable)")
    a = ap.parse_args()
    html = SITE.read_text(encoding="utf-8")
    before = html
    html = rates(html, a.apply)
    html = weather(html, a.apply)
    html = prices(html, a.set)
    if html != before:
        SITE.write_text(html, encoding="utf-8")
        print("\nSaved index.html. Check the site, then commit and push.")
    else:
        print("\nNothing changed." + ("" if a.apply else " Add --apply to write the new rates and temperatures."))


if __name__ == "__main__":
    sys.exit(main())
