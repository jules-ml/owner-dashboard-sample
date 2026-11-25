#!/usr/bin/env python3
"""Generate deliberately messy synthetic source exports for a fictional sample company.

Company: Lantern Key Pool & Air (FICTIONAL). Delray Beach / Palm Beach County, FL.
All names, phones, and emails are made up. Emails use the .invalid TLD.
As-of date: 2026-11-25. Activity window: 2026-01-06 through 2026-11-20.

This script writes only raw (messy) exports. It does not write clean tables.
"""
from __future__ import annotations

import csv
import random
from datetime import date, datetime, timedelta
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font

ROOT = Path(__file__).resolve().parent
RAW = ROOT / "data" / "raw"
SEED = 20261125
AS_OF = date(2026, 11, 25)
WINDOW_END = date(2026, 11, 20)

SERVICE_LINES = (
    "Weekly Pool Service",
    "Pool Repair",
    "HVAC Maintenance",
    "HVAC Repair",
    "Equipment Install",
)

# Target blended margins used when assigning job costs (direct labor + materials).
MARGINS = {
    "Weekly Pool Service": 0.58,
    "Pool Repair": 0.42,
    "HVAC Maintenance": 0.52,
    "HVAC Repair": 0.46,
    "Equipment Install": 0.28,
}

SERVICE_ALIASES = {
    "Weekly Pool Service": [
        "Weekly Pool Service",
        "Pool Cleaning",
        "pool clean",
        "Weekly Pool",
        "POOL SVC",
        "Pool Service - Monthly",
        "Pool Service",
    ],
    "Pool Repair": [
        "Pool Repair",
        "pool repairs",
        "Pump / Filter Repair",
        "Pool Equip Repair",
        "Pool Equipment Repair",
    ],
    "HVAC Maintenance": [
        "HVAC Maintenance",
        "AC Tune-Up",
        "hvac maint",
        "HVAC PMA",
        "A/C Maintenance",
    ],
    "HVAC Repair": [
        "HVAC Repair",
        "hvac-repair",
        "AC Service",
        "A/C Repair",
        "Air Conditioning Repair",
    ],
    "Equipment Install": [
        "Equipment Install",
        "Install",
        "AC Replacement",
        "Heater Install",
        "Equipment Installation",
    ],
}

TECHS = ("Luis M.", "Andre T.", "Keisha P.")

CUSTOMERS = [
    # Recurring weekly pool routes (billed monthly)
    dict(id="C001", name="Maria Gonzalez", aliases=["Maria Gonzalez", "Maria Gonzales", "M. Gonzalez", "Gonzalez, Maria", "MARIA GONZALEZ"], email="maria.gonzalez.sample@email.invalid", phone="5615550142", typ="residential", route=185.00),
    dict(id="C002", name="James O'Brien", aliases=["James O'Brien", "Jim OBrien", "James O Brien", "O'Brien, James", "J. O'Brien"], email="james.obrien.sample@email.invalid", phone="5615550188", typ="residential", route=195.00),
    dict(id="C003", name="Sarah Chen", aliases=["Sarah Chen", "sara chen", "Chen, Sarah", "S. Chen"], email="sarah.chen.sample@email.invalid", phone="5615550117", typ="residential", route=165.00),
    dict(id="C004", name="Robert Alvarez", aliases=["Robert Alvarez", "Bob Alvarez", "R. Alvarez", "Alvarez, Robert"], email="robert.alvarez.sample@email.invalid", phone="9545550133", typ="residential", route=175.00),
    dict(id="C005", name="Aisha Patel", aliases=["Aisha Patel", "Patel, Aisha", "Aisha  Patel"], email="aisha.patel.sample@email.invalid", phone="5615550164", typ="residential", route=180.00),
    dict(id="C006", name="Michael Thompson", aliases=["Michael Thompson", "Mike Thompson", "M Thompson", "Thompson, Michael"], email="mike.thompson.sample@email.invalid", phone="5615550190", typ="residential", route=170.00),
    dict(id="C007", name="Elena Rossi", aliases=["Elena Rossi", "Rossi, Elena", "ELENA ROSSI"], email="elena.rossi.sample@email.invalid", phone="9545550171", typ="residential", route=190.00),
    dict(id="C008", name="David Kim", aliases=["David Kim", "D. Kim", "Kim, David", "david kim"], email="david.kim.sample@email.invalid", phone="5615550108", typ="residential", route=160.00),
    dict(id="C009", name="Anthony Ruiz", aliases=["Anthony Ruiz", "Tony Ruiz", "Anthony  Ruiz", "Ruiz, Anthony"], email="anthony.ruiz.sample@email.invalid", phone="5615550156", typ="residential", route=200.00),
    dict(id="C010", name="Priya Nair", aliases=["Priya Nair", "PRIYA NAIR", "Nair, Priya", "P. Nair"], email="priya.nair.sample@email.invalid", phone="9545550122", typ="residential", route=175.00),
    dict(id="C011", name="Christopher Walsh", aliases=["Christopher Walsh", "Chris Walsh", "Walsh, Christopher", "C. Walsh"], email="chris.walsh.sample@email.invalid", phone="5615550181", typ="residential", route=185.00),
    dict(id="C012", name="Sofia Marino", aliases=["Sofia Marino", "Marino, Sofia", "S Marino"], email="sofia.marino.sample@email.invalid", phone="5615550149", typ="residential", route=205.00),
    # Residential, jobs only (no monthly route)
    dict(id="C013", name="Lauren Brooks", aliases=["Lauren Brooks", "L Brooks", "Brooks, Lauren", "lauren brooks"], email="lauren.brooks.sample@email.invalid", phone="5615550177", typ="residential", route=None),
    dict(id="C014", name="Diane Moreau", aliases=["Diane Moreau", "D. Moreau", "Moreau, Diane"], email="diane.moreau.sample@email.invalid", phone="9545550195", typ="residential", route=None),
    dict(id="C015", name="Hassan Farouk", aliases=["Hassan Farouk", "H. Farouk", "Farouk, Hassan"], email="hassan.farouk.sample@email.invalid", phone="5615550134", typ="residential", route=None),
    dict(id="C016", name="Nicole Bennett", aliases=["Nicole Bennett", "N. Bennett", "Bennett, Nicole", "NICOLE BENNETT"], email="nicole.bennett.sample@email.invalid", phone="5615550169", typ="residential", route=None),
    dict(id="C017", name="Thomas Grant", aliases=["Thomas Grant", "Tom Grant", "Grant, Thomas", "T. Grant"], email="tom.grant.sample@email.invalid", phone="9545550104", typ="residential", route=None),
    dict(id="C018", name="Kevin Doyle", aliases=["Kevin Doyle", "K Doyle", "Doyle, Kevin"], email="kevin.doyle.sample@email.invalid", phone="5615550128", typ="residential", route=None),
    dict(id="C019", name="Jasmine Cole", aliases=["Jasmine Cole", "J. Cole", "Cole, Jasmine"], email="jasmine.cole.sample@email.invalid", phone="5615550152", typ="residential", route=None),
    dict(id="C020", name="Patrick Nguyen", aliases=["Patrick Nguyen", "Pat Nguyen", "Nguyen, Patrick", "P. Nguyen"], email="patrick.nguyen.sample@email.invalid", phone="9545550186", typ="residential", route=None),
    # Commercial / HOA
    dict(id="C021", name="Ocean Palms HOA", aliases=["Ocean Palms HOA", "Ocean Palms H.O.A.", "Ocean Palms Hoa Inc", "Ocean Palms Hoa"], email="office.oceanpalms.sample@email.invalid", phone="5615550201", typ="commercial", route=420.00),
    dict(id="C022", name="Riviera Court Condos", aliases=["Riviera Court Condos", "Riviera Court Condo Assoc", "Riviera Court Condos Inc", "Riviera Court Condo"], email="manager.rivieracourt.sample@email.invalid", phone="5615550214", typ="commercial", route=380.00),
    dict(id="C023", name="Cypress Walk Association", aliases=["Cypress Walk Association", "Cypress Walk Assn", "Cypress Walk HOA", "Cypress Walk Assoc"], email="board.cypresswalk.sample@email.invalid", phone="5615550226", typ="commercial", route=350.00),
    dict(id="C024", name="Breeze Bay Club", aliases=["Breeze Bay Club", "Breeze Bay Club LLC", "BreezeBay Club", "Breeze Bay Club, LLC"], email="ops.breezebay.sample@email.invalid", phone="9545550239", typ="commercial", route=None),
    dict(id="C025", name="Harbor Oaks Property Management", aliases=["Harbor Oaks Property Management", "Harbor Oaks PM", "Harbor Oaks Property Mgmt", "Harbor Oaks Property Mgmt."], email="ap.harboroaks.sample@email.invalid", phone="5615550245", typ="commercial", route=None),
]


def daterange_months():
    return [date(2026, m, 1) for m in range(1, 12)]


def excel_serial(d: date) -> int:
    return (datetime(d.year, d.month, d.day) - datetime(1899, 12, 30)).days


def mess_date(rng: random.Random, d: date, allow_blank: bool = True) -> object:
    """Return a messy date representation."""
    roll = rng.random()
    if allow_blank and roll < 0.025:
        return rng.choice(["", "N/A", "n/a", None])
    formats = [
        lambda: d.strftime("%Y-%m-%d"),
        lambda: d.strftime("%m/%d/%Y"),
        lambda: d.strftime("%-m/%-d/%y"),
        lambda: d.strftime("%d-%b-%Y"),
        lambda: d.strftime("%B %d, %Y"),
        lambda: d.strftime("%Y/%m/%d"),
        lambda: d.strftime("%m.%d.%Y"),
        lambda: excel_serial(d),
        lambda: d.strftime("%m/%d/%Y"),
        lambda: d.strftime("%-m/%-d/%Y"),
    ]
    # A few European-style day-first values when day > 12 so they are unambiguous.
    if d.day > 12 and rng.random() < 0.12:
        return d.strftime("%d/%m/%Y")
    return rng.choice(formats)()


def mess_money(rng: random.Random, amount: float, allow_blank: bool = False) -> object:
    if allow_blank and rng.random() < 0.015:
        return rng.choice(["", "n/a", "N/A", None])
    if amount < 0:
        return rng.choice([
            f"({abs(amount):,.2f})",
            f"-${abs(amount):,.2f}",
            f"-{abs(amount):.2f}",
        ])
    styles = [
        lambda: f"${amount:,.2f}",
        lambda: f"{amount:.2f}",
        lambda: f"{amount:,.2f}",
        lambda: f"USD {amount:,.2f}",
        lambda: f"{amount:,.2f} USD",
        lambda: f"{amount:.2f}$",
        lambda: f"  ${amount:,.2f} ",
        lambda: f"${amount:,.2f} ",
        lambda: f"{amount:.0f}" if amount == int(amount) else f"{amount:.2f}",
        lambda: f"$ {amount:,.2f}",
    ]
    # A few European thousands/decimal formats.
    if rng.random() < 0.06:
        whole = f"{int(round(amount * 100)):d}"
        # build 1.250,00 from cents
        cents = int(round(amount * 100))
        core = f"{cents // 100:,}".replace(",", ".") + f",{cents % 100:02d}"
        return core
    return rng.choice(styles)()


def mess_invoice_id(rng: random.Random, n: int) -> str:
    return rng.choice([
        f"INV-{n}",
        f"inv{n}",
        f"#{n}",
        f"{n}",
        f"INV{n}",
        f"INV-{n} ",
        f"inv-{n}",
        f"Inv {n}",
    ])


def mess_phone(rng: random.Random, digits: str, allow_blank: bool = True) -> str:
    if allow_blank and rng.random() < 0.12:
        return rng.choice(["", "n/a", "-"])
    a, b, c = digits[:3], digits[3:6], digits[6:]
    return rng.choice([
        f"({a}) {b}-{c}",
        f"{a}-{b}-{c}",
        f"{a}.{b}.{c}",
        f"{a}{b}{c}",
        f"+1 {a} {b} {c}",
        f"{a}-{b}-{c}",
    ])


def pick_alias(rng: random.Random, customer: dict) -> str:
    return rng.choice(customer["aliases"])


def pick_service_alias(rng: random.Random, line: str) -> str:
    return rng.choice(SERVICE_ALIASES[line])


def due_for(inv_date: date, typ: str) -> date:
    if typ == "commercial":
        return inv_date + timedelta(days=30)
    return inv_date + timedelta(days=15)


def add_months(d: date, months: int) -> date:
    m = d.month - 1 + months
    y = d.year + m // 12
    m = m % 12 + 1
    day = min(d.day, [31, 29 if y % 4 == 0 else 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][m - 1])
    return date(y, m, day)


def main() -> None:
    rng = random.Random(SEED)
    RAW.mkdir(parents=True, exist_ok=True)

    invoices: list[dict] = []
    payments: list[dict] = []
    jobs: list[dict] = []
    inv_num = 1001
    job_num = 4001

    def new_invoice(customer, inv_date, line, amount, status_plan, notes=""):
        nonlocal inv_num
        inv = {
            "inv_n": inv_num,
            "customer": customer,
            "date": inv_date,
            "due": due_for(inv_date, customer["typ"]),
            "line": line,
            "amount": round(amount, 2),
            "tax": round(amount * 0.07, 2),
            "status_plan": status_plan,
            "notes": notes,
        }
        inv_num += 1
        invoices.append(inv)
        return inv

    def new_job(customer, job_date, line, invoice, hours=None):
        nonlocal job_num
        margin = MARGINS[line]
        # Direct cost with a little noise around the target margin.
        noise = rng.uniform(-0.04, 0.04)
        cost = round(invoice["amount"] * (1 - (margin + noise)), 2)
        labor_share = 0.62 if line != "Equipment Install" else 0.35
        labor = round(cost * labor_share, 2)
        materials = round(cost - labor, 2)
        if hours is None:
            hours = round(max(1.0, labor / rng.uniform(42, 58)), 1)
        job = {
            "job_n": job_num,
            "customer": customer,
            "date": job_date,
            "line": line,
            "invoice": invoice,
            "labor": labor,
            "materials": materials,
            "tech": rng.choice(TECHS),
            "hours": hours,
            "status": "Completed",
        }
        job_num += 1
        jobs.append(job)
        return job

    # --- Recurring monthly pool service, Jan–Nov 2026 ---
    route_customers = [c for c in CUSTOMERS if c["route"]]
    for month_start in daterange_months():
        for c in route_customers:
            bill_day = 1 + (int(c["id"][1:]) % 5)
            inv_date = date(month_start.year, month_start.month, bill_day)
            if inv_date > WINDOW_END:
                continue
            amount = c["route"]
            # Mid-year rate bump for some routes.
            if month_start.month >= 6 and int(c["id"][1:]) % 3 == 0:
                amount = round(amount * 1.06, 2)
            inv = new_invoice(c, inv_date, "Weekly Pool Service", amount, status_plan="auto")
            # One summarized route job per monthly invoice (chemicals + tech time).
            new_job(c, inv_date + timedelta(days=1), "Weekly Pool Service", inv, hours=rng.choice([3.0, 3.5, 4.0, 5.0, 6.0] if c["typ"] == "commercial" else [1.5, 2.0, 2.5]))

    # --- Discrete jobs through the year ---
    discrete = [
        # (month, day, customer_id, line, amount, notes)
        (1, 9, "C013", "HVAC Repair", 890, "Capacitor and contactor"),
        (1, 14, "C015", "Pool Repair", 640, "Pump basket and lid"),
        (1, 21, "C024", "HVAC Maintenance", 1800, "Clubhouse plus four units"),
        (1, 27, "C017", "Equipment Install", 6400, "3-ton air handler swap"),
        (1, 30, "C018", "Pool Repair", 425, "Filter grid pack"),
        (2, 5, "C016", "HVAC Repair", 1120, "Blower motor"),
        (2, 11, "C019", "Pool Repair", 780, "Salt cell replacement"),
        (2, 18, "C014", "HVAC Maintenance", 220, "Spring tune-up"),
        (2, 20, "C020", "HVAC Repair", 1480, "Evaporator coil leak repair"),
        (2, 26, "C013", "Pool Repair", 540, "Heater igniter"),
        (3, 4, "C025", "HVAC Maintenance", 2400, "Managed properties, March PMA"),
        (3, 7, "C007", "Pool Repair", 1260, "Variable-speed pump"),
        (3, 12, "C015", "HVAC Repair", 960, "Capacitor, fan motor"),
        (3, 16, "C021", "Equipment Install", 8200, "Pool heater install"),
        (3, 19, "C024", "HVAC Repair", 2140, "Compressor start components"),
        (3, 25, "C017", "HVAC Maintenance", 195, "Post-install checkup"),
        (3, 28, "C012", "Pool Repair", 890, "Leak detection and patch"),
        (4, 2, "C016", "HVAC Maintenance", 210, "Spring PMA"),
        (4, 8, "C018", "Pool Repair", 1560, "Pool heater gas valve"),
        (4, 10, "C014", "HVAC Repair", 1880, "TXV and drier"),
        (4, 15, "C023", "HVAC Maintenance", 2100, "Association clubhouse and gym"),
        (4, 18, "C020", "Equipment Install", 4100, "Package unit, small condo"),
        (4, 22, "C009", "Pool Repair", 470, "Timer and relay"),
        (4, 28, "C019", "HVAC Repair", 740, "Capacitor, condensate flush"),
        (5, 5, "C013", "HVAC Repair", 2340, "Compressor hard start plus recharge"),
        (5, 8, "C025", "Pool Repair", 1680, "Commercial pump rebuild"),
        (5, 13, "C015", "Equipment Install", 7600, "4-ton split system"),
        (5, 19, "C006", "Pool Repair", 610, "Filter tank clamp"),
        (5, 21, "C016", "HVAC Repair", 1290, "Control board"),
        (5, 27, "C024", "HVAC Repair", 1860, "Two RTU service calls"),
        (6, 3, "C014", "HVAC Repair", 1540, "Fan motor and capacitor"),
        (6, 6, "C017", "HVAC Repair", 980, "Frozen coil, low charge"),
        (6, 11, "C020", "HVAC Repair", 2110, "Condenser fan and run cap"),
        (6, 14, "C025", "Equipment Install", 8900, "Two package units, building B"),
        (6, 18, "C013", "HVAC Repair", 670, "Thermostat and float switch"),
        (6, 23, "C018", "Pool Repair", 1320, "Salt system and cell"),
        (6, 27, "C019", "HVAC Repair", 2450, "Compressor replacement labor"),
        (7, 2, "C016", "HVAC Repair", 1760, "Shorted compressor diagnosis + start kit"),
        (7, 8, "C014", "HVAC Repair", 2040, "Evaporator leak, coil treatment"),
        (7, 11, "C024", "HVAC Repair", 2680, "Club RTU repair"),
        (7, 16, "C021", "Pool Repair", 2140, "Heater heat-exchanger work"),
        (7, 22, "C025", "Equipment Install", 4500, "Air handler, building C (balance billed later)"),
        (7, 24, "C008", "Pool Repair", 390, "Impeller"),
        (7, 29, "C017", "HVAC Repair", 1180, "Blower and capacitor"),
        (8, 5, "C013", "HVAC Repair", 1420, "Contactor, capacitor, hard start"),
        (8, 8, "C020", "HVAC Repair", 1880, "Low voltage short, transformer"),
        (8, 13, "C015", "HVAC Repair", 2310, "TXV and recharge"),
        (8, 19, "C016", "HVAC Repair", 960, "Drain and condensate pump"),
        (8, 21, "C014", "HVAC Repair", 2740, "Compressor changeout"),
        (8, 26, "C019", "Pool Repair", 820, "Automation panel"),
        (8, 28, "C018", "HVAC Repair", 1540, "Fan motor"),
        (9, 3, "C024", "HVAC Maintenance", 1800, "Fall PMA, club"),
        (9, 9, "C017", "Equipment Install", 6800, "Heat-pump swap"),
        (9, 12, "C013", "HVAC Repair", 740, "Capacitor and clean"),
        (9, 16, "C023", "HVAC Repair", 1920, "Gym unit board and sensors"),
        (9, 18, "C006", "Pool Repair", 1140, "Pump and union leaks"),
        (9, 23, "C025", "HVAC Maintenance", 2400, "Fall PMA, managed properties"),
        (9, 29, "C020", "HVAC Repair", 860, "Float switch, drain"),
        (10, 2, "C016", "HVAC Maintenance", 210, "Fall tune-up"),
        (10, 7, "C014", "HVAC Maintenance", 220, "Fall tune-up"),
        (10, 9, "C021", "Equipment Install", 9200, "Variable-speed pump + automation (deposit billed)"),
        (10, 14, "C015", "Pool Repair", 980, "Heater bypass and valves"),
        (10, 17, "C024", "HVAC Repair", 1640, "Club RTU sensors"),
        (10, 21, "C019", "Equipment Install", 0, "placeholder"),  # replaced below if 0
        (10, 22, "C017", "HVAC Repair", 1320, "Capacitor and contactors"),
        (10, 28, "C013", "HVAC Maintenance", 195, "Fall tune-up"),
        (11, 4, "C018", "Pool Repair", 610, "Filter and laterals"),
        (11, 6, "C016", "HVAC Repair", 890, "No heat, igniter and flame sensor"),
        (11, 10, "C025", "HVAC Repair", 1480, "Building A no-cool"),
        (11, 12, "C014", "Equipment Install", 4800, "Air handler only"),
        (11, 14, "C020", "HVAC Maintenance", 230, "Fall tune-up, late"),
        (11, 18, "C019", "Pool Repair", 540, "Salt cell, under warranty labor"),
        (11, 19, "C007", "HVAC Repair", 760, "Capacitor"),
    ]
    # Fix the placeholder if any
    discrete = [row for row in discrete if row[4] != 0]
    # Add a mid-October install that was meant to be in the list
    discrete.append((10, 21, "C019", "HVAC Maintenance", 210, "Fall tune-up"))

    by_id = {c["id"]: c for c in CUSTOMERS}
    for month, day, cid, line, amount, notes in discrete:
        inv_date = date(2026, month, day)
        if inv_date > WINDOW_END:
            continue
        c = by_id[cid]
        inv = new_invoice(c, inv_date, line, amount, status_plan="auto", notes=notes)
        job_date = inv_date - timedelta(days=rng.choice([0, 0, 1, 2]))
        new_job(c, job_date, line, inv)

    # --- Planned payment behavior (truth), including planted AR problems ---
    # Default: residential paid 4–18 days after invoice; commercial 20–40 days.
    # Overrides below create the chase-first list as of 2026-11-25.
    def default_pay_date(inv) -> date | None:
        typ = inv["customer"]["typ"]
        lag = rng.randint(20, 40) if typ == "commercial" else rng.randint(4, 18)
        pay = inv["date"] + timedelta(days=lag)
        if pay > AS_OF:
            return None
        return pay

    # Planted overdue / open invoices (by inv notes, customer, date, amount).
    # Matched after generation by (customer id, date, amount).
    planted_open = {
        # Harbor Oaks install balance from July 22 — 118 days past due on Nov 25 (due Aug 21)
        ("C025", date(2026, 7, 22), 4500.00): "unpaid",
        # Ocean Palms pool repair July 16 — 96 days past due (due Aug 15 commercial 30-day)
        ("C021", date(2026, 7, 16), 2140.00): "unpaid",
        # Ocean Palms: leave a second older invoice partially unpaid? October install deposit unpaid portion
        ("C021", date(2026, 10, 9), 9200.00): "partial_2140",  # paid 7060, open 2140 → wait, I wanted ~4280 total for Ocean Palms
        # James O'Brien: use a special planted repair (added below if missing)
        ("C002", date(2026, 9, 4), 2140.00): "unpaid",
        # Riviera Court HVAC maint April was paid; plant September? Let's plant a September invoice
        ("C022", date(2026, 9, 8), 1875.00): "unpaid",
        # Lauren Brooks Aug 21 compressor — actually I have Aug 5 for C013. Use Aug 5? That's HVAC 1420.
        # Want $890: January 9 invoice might have been paid. Plant October leftover? Use May 5? No.
        # January 9 is $890 — if left unpaid it would be 300+ days. Plant a new Sept 28 invoice.
        ("C013", date(2026, 9, 28), 890.00): "unpaid",
        # Cypress Walk repair
        ("C023", date(2026, 10, 3), 1240.00): "unpaid",
        # Kevin Doyle filter
        ("C018", date(2026, 10, 8), 425.00): "unpaid",
        # Diane Moreau — Nov 3 invoice approaching 30? 22 days if dated Nov 3 and net-15 due Nov 18 → 7 days past due.
        # Want 22 days past due: invoice Oct 19, due Nov 3 (residential 15-day), 22 days past due on Nov 25.
        ("C014", date(2026, 10, 19), 1150.00): "unpaid",
    }

    # Add planted invoices that are not already in the discrete list.
    extras = [
        (9, 4, "C002", "Pool Repair", 2140.00, "Pump and control board — waiting on insurance"),
        (9, 8, "C022", "HVAC Maintenance", 1875.00, "Fall common-area PMA"),
        (9, 28, "C013", "HVAC Repair", 890.00, "Capacitor and service call, no response to follow-up"),
        (10, 3, "C023", "Pool Repair", 1240.00, "Commercial pump seal and bearings"),
        (10, 8, "C018", "Pool Repair", 425.00, "Filter grid pack"),
        (10, 19, "C014", "HVAC Repair", 1150.00, "Fan motor and capacitor"),
    ]
    existing_keys = {(inv["customer"]["id"], inv["date"], inv["amount"]) for inv in invoices}
    for month, day, cid, line, amount, notes in extras:
        key = (cid, date(2026, month, day), amount)
        if key in existing_keys:
            continue
        c = by_id[cid]
        inv_date = date(2026, month, day)
        inv = new_invoice(c, inv_date, line, amount, status_plan="auto", notes=notes)
        new_job(c, inv_date, line, inv)
        existing_keys.add(key)

    # Assign payments from status plans
    for inv in invoices:
        key = (inv["customer"]["id"], inv["date"], inv["amount"])
        plan = planted_open.get(key, "auto")
        if plan == "unpaid":
            inv["status_plan"] = "Open"
            continue
        if plan == "partial_2140":
            inv["status_plan"] = "Partial"
            pay_date = inv["date"] + timedelta(days=18)
            payments.append({
                "invoice": inv,
                "date": pay_date,
                "amount": 7060.00,
                "method": "ACH",
            })
            continue
        # Auto: unpaid if the default pay date would fall after AS_OF (recent invoices).
        pay_date = default_pay_date(inv)
        if pay_date is None:
            inv["status_plan"] = "Open"
            continue
        # A few partials and a couple of odd overpays for the cleaner to flag.
        roll = rng.random()
        if roll < 0.04 and inv["amount"] > 400 and pay_date <= AS_OF:
            inv["status_plan"] = "Partial"
            payments.append({
                "invoice": inv,
                "date": pay_date,
                "amount": round(inv["amount"] * rng.uniform(0.35, 0.7), 2),
                "method": rng.choice(["Check", "ACH", "Visa", "Cash"]),
            })
        else:
            inv["status_plan"] = "Paid"
            method = rng.choice(["Check", "Check", "ACH", "ACH", "Visa", "Visa", "Mastercard", "Cash"])
            payments.append({
                "invoice": inv,
                "date": pay_date,
                "amount": inv["amount"],
                "method": method,
            })

    # One deliberate overpayment for the quality log (flag, do not auto-correct).
    # Pick a paid mid-year residential invoice and add $40.
    for p in payments:
        if p["invoice"]["customer"]["typ"] == "residential" and p["invoice"]["date"].month == 5 and p["amount"] == p["invoice"]["amount"]:
            p["amount"] = round(p["amount"] + 40.00, 2)
            p["invoice"]["notes"] = (p["invoice"]["notes"] + " | overpay $40 left as-is").strip(" |")
            break

    # One credit memo (negative invoice) for a warranty adjustment.
    credit_cust = by_id["C008"]
    credit_date = date(2026, 8, 14)
    credit = new_invoice(credit_cust, credit_date, "Pool Repair", -180.00, status_plan="Paid", notes="Warranty credit, impeller job")
    payments.append({
        "invoice": credit,
        "date": credit_date,
        "amount": -180.00,
        "method": "Credit memo",
    })

    # --- Operating expenses (overhead only; job labor/materials are separate) ---
    vendors = {
        "fuel": ("Gulf Sample Fuel Card", "Fuel"),
        "insurance": ("Sample Mutual Insurance", "Insurance"),
        "truck": ("Southeast Sample Truck Lease", "Vehicle lease"),
        "rent": ("Atlantic Sample Plaza LLC", "Shop rent"),
        "software": ("Routebook Software (sample)", "Software"),
        "uniforms": ("Coastal Sample Uniforms", "Uniforms"),
        "waste": ("Sample Dump & Disposal", "Waste disposal"),
        "misc": ("Office Supply Sample Co.", "Office"),
    }
    expenses: list[dict] = []
    exp_id = 7001
    for month_start in daterange_months():
        m = month_start.month
        # Fixed monthly
        fixed = [
            ("insurance", date(2026, m, 1), 1850.00),
            ("truck", date(2026, m, 1), 1520.00),
            ("rent", date(2026, m, 1), 2400.00),
            ("software", date(2026, m, 3), 189.00),
        ]
        fuel = round(1100 + (220 if m in (6, 7, 8, 9) else 0) + rng.uniform(-80, 80), 2)
        fuel_date = date(2026, m, rng.randint(6, 12))
        waste = round(rng.uniform(85, 140), 2)
        waste_date = date(2026, m, rng.randint(14, 20))
        if m in (3, 9):
            uniforms = 340.00
            uni_date = date(2026, m, 8)
        else:
            uniforms = None
            uni_date = None
        misc = round(rng.uniform(60, 210), 2)
        misc_date = date(2026, m, rng.randint(10, 24))
        rows = list(fixed) + [("fuel", fuel_date, fuel), ("waste", waste_date, waste), ("misc", misc_date, misc)]
        if uniforms:
            rows.append(("uniforms", uni_date, uniforms))
        for key, dt, amount in rows:
            if dt > WINDOW_END:
                continue
            vendor, category = vendors[key]
            expenses.append({
                "exp_n": exp_id,
                "date": dt,
                "vendor": vendor,
                "category": category,
                "amount": amount,
                "method": "ACH" if key in ("insurance", "truck", "rent", "software") else rng.choice(["Visa", "ACH", "Check"]),
                "memo": category,
            })
            exp_id += 1

    # --- CRM leads ---
    lead_sources = ["Google", "Referral", "Yard Sign", "Nextdoor", "Angi", "Repeat Customer", "Facebook"]
    lead_people = [
        ("Owen Bradley", "New"),
        ("Camila Ortiz", "Contacted"),
        ("Ryan Phelps", "Quoted"),
        ("Hannah Yi", "Lost"),
        ("Marcus Bell", "Won"),
        ("Tina Vasquez", "Quoted"),
        ("Jonah Green", "Contacted"),
        ("Vera Hoffman", "Lost"),
        ("Samuel Ortiz", "New"),
        ("Katie Brooks", "Won"),
        ("Leo Abramson", "Contacted"),
        ("Nadia Shakir", "Quoted"),
        ("Peter Lang", "Lost"),
        ("Willa Chen", "New"),
        ("Omar Haddad", "Contacted"),
        ("Grace Little", "Lost"),
        ("Henry Cole", "Quoted"),
        ("Amelia Soto", "Won"),
        ("Brian Duffy", "Contacted"),
        ("Chloe Park", "New"),
        ("Dante Rivers", "Lost"),
        ("Ivy Moretti", "Quoted"),
        ("Quinn Barker", "Contacted"),
        ("Rosa Flint", "New"),
        ("Steven Cho", "Lost"),
        ("Uma Patel", "Quoted"),
        ("Felix Ward", "Contacted"),
        ("Yara Solano", "New"),
        ("Nate Crowley", "Lost"),
        ("Paula Simmons", "Quoted"),
        ("Ricardo Vega", "Won"),
        ("Sasha Quinn", "Contacted"),
        ("Theo March", "New"),
        ("Blair Jenkins", "Lost"),
        ("Connie Zhu", "Quoted"),
        ("Drew Alston", "Contacted"),
        ("Eden Clark", "New"),
        ("Farah Noor", "Quoted"),
        ("Gabe Lind", "Lost"),
        ("Hailey Frost", "Contacted"),
        ("Ian Brooks", "New"),
        ("Jules Marino", "Quoted"),
        ("Kara Singh", "Lost"),
        ("Miles Ortega", "Contacted"),
        ("Nina Capello", "New"),
    ]
    leads = []
    lead_n = 2001
    start = date(2026, 1, 8)
    for i, (name, status) in enumerate(lead_people):
        created = start + timedelta(days=int(i * (316 / len(lead_people))) + rng.randint(-3, 3))
        created = max(date(2026, 1, 6), min(created, WINDOW_END))
        est = rng.choice([220, 350, 480, 650, 890, 1200, 1800, 2400, 4500, 6800, 8200])
        source = lead_sources[i % len(lead_sources)]
        leads.append({
            "lead_n": lead_n,
            "created": created,
            "name": name,
            "source": source,
            "status": status,
            "est": float(est),
            "notes": "",
        })
        lead_n += 1
    # A few leads that became known customers (won, matching names loosely / same people)
    leads.append({
        "lead_n": lead_n,
        "created": date(2026, 1, 10),
        "name": "Lauren Brooks",
        "source": "Google",
        "status": "Won",
        "est": 900.00,
        "notes": "Became a customer (sample)",
    })

    # ========== Messy invoice+payment export (single CSV, denormalized) ==========
    # Some invoices with multiple payment rows; some exact duplicate rows; messy formats.
    payments_by_inv = {}
    for p in payments:
        payments_by_inv.setdefault(p["invoice"]["inv_n"], []).append(p)

    inv_rows = []
    methods_messy = {
        "Check": ["Check", "check", "CHK", "Check "],
        "ACH": ["ACH", "ach", "Bank transfer", "ACH "],
        "Visa": ["Visa", "visa ****1234", "VISA", "Card"],
        "Mastercard": ["Mastercard", "MC", "mastercard"],
        "Cash": ["Cash", "cash", "CASH"],
        "Credit memo": ["Credit memo", "CREDIT", "credit memo"],
    }

    # Invoices that will be emitted twice as exact duplicate rows.
    dup_targets = {invoices[i]["inv_n"] for i in (3, 18, 42, 77) if i < len(invoices)}
    # A few invoices with unparseable amounts or missing dates (flagged, not silently dropped if we can still identify them).
    bad_amount_ids = set()
    bad_date_ids = set()
    # Pick identifiable mid-list invoices that are paid and otherwise ordinary.
    ordinary = [inv for inv in invoices if inv["status_plan"] == "Paid" and inv["amount"] > 0 and inv["date"].month in (3, 4)]
    for inv in ordinary[:3]:
        bad_amount_ids.add(inv["inv_n"])
    for inv in ordinary[3:5]:
        bad_date_ids.add(inv["inv_n"])
    # One typo year 2027
    typo_year_id = ordinary[5]["inv_n"] if len(ordinary) > 5 else None

    def status_out(inv, has_payment_on_row):
        plan = inv["status_plan"]
        if not has_payment_on_row and plan == "Paid":
            # sometimes the status is blank or inconsistent on a payment-less view
            return rng.choice(["Paid", "paid", "PAID", "Open", ""])
        return {
            "Paid": rng.choice(["Paid", "paid", "PAID", "Paid "]),
            "Partial": rng.choice(["Partial", "partial", "PART PAID"]),
            "Open": rng.choice(["Open", "open", "OVERDUE", "overdue", ""]),
        }[plan]

    for inv in invoices:
        pays = payments_by_inv.get(inv["inv_n"], [])
        rows_for_inv = pays or [None]
        for p in rows_for_inv:
            # Occasional blank email
            email = inv["customer"]["email"]
            if rng.random() < 0.08:
                email = rng.choice(["", " ", email.upper(), email.replace("@", "@@")])
            amount_out = mess_money(rng, inv["amount"])
            date_out = mess_date(rng, inv["date"], allow_blank=False)
            if inv["inv_n"] in bad_amount_ids:
                amount_out = rng.choice(["#REF!", "??", "tbd", "see notes"])
            if inv["inv_n"] in bad_date_ids:
                date_out = rng.choice(["", "N/A", "00/00/0000", "not recorded"])
            if typo_year_id and inv["inv_n"] == typo_year_id:
                date_out = inv["date"].replace(year=2027).strftime("%m/%d/%Y")
            row = {
                "Invoice #": mess_invoice_id(rng, inv["inv_n"]),
                "Inv Date": date_out,
                "Due Date": mess_date(rng, inv["due"], allow_blank=True),
                "Customer Name": pick_alias(rng, inv["customer"]),
                "Email": email,
                "Phone": mess_phone(rng, inv["customer"]["phone"]),
                "Service": pick_service_alias(rng, inv["line"]),
                "Amount": amount_out,
                "Tax": mess_money(rng, inv["tax"]) if rng.random() > 0.1 else "",
                "Status": status_out(inv, p is not None),
                "Payment Date": mess_date(rng, p["date"], allow_blank=False) if p else "",
                "Payment Amt": mess_money(rng, p["amount"]) if p else "",
                "Method": rng.choice(methods_messy.get(p["method"], [p["method"]])) if p else "",
                "Notes": inv["notes"],
            }
            inv_rows.append(row)
            if inv["inv_n"] in dup_targets and p is rows_for_inv[0]:
                inv_rows.append(dict(row))  # exact duplicate

    # Shuffle so the export doesn't look pre-sorted (but keep a little date clustering).
    # Actually a messy export might be roughly chronological with dups inserted. Shuffle lightly in chunks.
    rng.shuffle(inv_rows)

    inv_path = RAW / "invoices_payments.csv"
    fieldnames = ["Invoice #", "Inv Date", "Due Date", "Customer Name", "Email", "Phone", "Service", "Amount", "Tax", "Status", "Payment Date", "Payment Amt", "Method", "Notes"]
    with inv_path.open("w", newline="", encoding="utf-8") as f:
        # Deliberately messy: a preamble line some exports include, which the cleaner must skip.
        # We'll write a standard CSV (cleaner can handle utf-8) — preamble would break csv readers
        # unless the cleaner skips it. I'll put a junk first line.
        f.write("QuickBooks Online invoice export — Lantern Key Pool & Air — 11/20/2026\n")
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(inv_rows)

    # ========== Messy jobs + expenses workbook ==========
    wb = Workbook()
    # Jobs sheet with title rows before headers
    ws = wb.active
    ws.title = "Jobs Export"
    ws["A1"] = "Job costing export — pulled 11/20/2026 — SAMPLE DATA"
    ws["A1"].font = Font(bold=True)
    ws["A2"] = "Lantern Key Pool & Air (fictional)"
    # headers on row 4
    job_headers = ["Job Number", "Date", "Customer", "Type", "Invoice Ref", "Labor", "Materials", "Tech", "Hours", "Job Status", "Notes"]
    for col, h in enumerate(job_headers, 1):
        ws.cell(4, col, h).font = Font(bold=True)

    # Two jobs with no matching customer (flagged orphans)
    orphan_jobs = [
        dict(job_n=job_num, date=date(2026, 6, 19), customer_name="Unknown walk-in", line="Pool Repair", invoice=None, labor=180.00, materials=95.00, tech="Luis M.", hours=3.0, status="Completed", notes="Name not in customer list"),
        dict(job_n=job_num + 1, date=date(2026, 9, 15), customer_name="do not use - test row", line="HVAC Maintenance", invoice=None, labor=0, materials=0, tech="", hours=0, status="", notes="Looks like a test row"),
    ]
    job_num += 2

    # One cancelled job (should not count as cost against revenue)
    cancelled = dict(
        job_n=job_num,
        date=date(2026, 4, 19),
        customer=by_id["C016"],
        line="HVAC Repair",
        invoice=None,
        labor=0,
        materials=0,
        tech="Andre T.",
        hours=0,
        status="Cancelled",
        notes="Customer cancelled, unit under warranty elsewhere",
    )
    job_num += 1

    r = 5
    all_jobs = list(jobs) + orphan_jobs + [cancelled]
    rng.shuffle(all_jobs)
    for job in all_jobs:
        is_orphan = "customer_name" in job
        cust_name = job["customer_name"] if is_orphan else pick_alias(rng, job["customer"])
        inv_ref = ""
        if job.get("invoice"):
            inv_ref = rng.choice([
                f"INV-{job['invoice']['inv_n']}",
                f"{job['invoice']['inv_n']}",
                f"#{job['invoice']['inv_n']}",
                f"inv{job['invoice']['inv_n']}",
            ])
        status = job["status"]
        if status:
            status = rng.choice([status, status.lower(), status.upper(), status])
        ws.cell(r, 1, rng.choice([f"JOB-{job['job_n']}", f"job {job['job_n']}", str(job["job_n"]), f"J-{job['job_n']}"]))
        ws.cell(r, 2, mess_date(rng, job["date"], allow_blank=False))
        ws.cell(r, 3, cust_name)
        ws.cell(r, 4, pick_service_alias(rng, job["line"]) if job.get("line") in SERVICE_ALIASES else job.get("line", ""))
        ws.cell(r, 5, inv_ref)
        ws.cell(r, 6, mess_money(rng, job["labor"]) if job["labor"] or job["labor"] == 0 else "")
        ws.cell(r, 7, mess_money(rng, job["materials"]))
        ws.cell(r, 8, job["tech"])
        ws.cell(r, 9, job["hours"] if job["hours"] else "")
        ws.cell(r, 10, status)
        ws.cell(r, 11, job.get("notes", ""))
        r += 1

    # Expenses sheet, also messy
    we = wb.create_sheet("Expenses")
    we["A1"] = "Bank / card expenses — SAMPLE — do not use as live books"
    exp_headers = ["Date", "Vendor", "Category", "Amount", "Payment Method", "Memo", "Ref"]
    for col, h in enumerate(exp_headers, 1):
        we.cell(3, col, h).font = Font(bold=True)
    # Duplicate one expense row
    exp_rows = list(expenses)
    if exp_rows:
        exp_rows.append(dict(exp_rows[4]))
    rr = 4
    for exp in exp_rows:
        we.cell(rr, 1, mess_date(rng, exp["date"], allow_blank=False))
        vendor = exp["vendor"]
        if rng.random() < 0.15:
            vendor = vendor.upper() if rng.random() < 0.5 else vendor + " "
        we.cell(rr, 2, vendor)
        we.cell(rr, 3, exp["category"] if rng.random() > 0.08 else exp["category"].lower())
        we.cell(rr, 4, mess_money(rng, exp["amount"]))
        we.cell(rr, 5, exp["method"] if rng.random() > 0.1 else exp["method"].lower())
        we.cell(rr, 6, exp["memo"])
        we.cell(rr, 7, f"E-{exp['exp_n']}")
        rr += 1
    wb.save(RAW / "jobs_expenses.xlsx")

    # ========== Messy CRM leads ==========
    lead_path = RAW / "crm_leads.csv"
    lead_fieldnames = ["Lead ID", "Created On", "Full Name", "Src", "Lead Status", "Est. Value", "Notes"]
    with lead_path.open("w", newline="", encoding="utf-8") as f:
        f.write("CRM export (sample) — Shoreline? wait, Lantern Key Pool & Air — pulled 11/20/2026\n")
        writer = csv.DictWriter(f, fieldnames=lead_fieldnames)
        writer.writeheader()
        for lead in leads:
            name = lead["name"]
            if rng.random() < 0.2:
                parts = name.split()
                if len(parts) == 2:
                    name = rng.choice([f"{parts[1]}, {parts[0]}", name.upper(), name.lower(), name])
            src = lead["source"]
            if rng.random() < 0.15:
                src = rng.choice([src.lower(), src.upper(), src + " ", "fb" if src == "Facebook" else src])
            status = lead["status"]
            if rng.random() < 0.2:
                status = rng.choice([status.lower(), status.upper(), "WON" if status == "Won" else status, ""])
            writer.writerow({
                "Lead ID": rng.choice([f"L-{lead['lead_n']}", f"lead{lead['lead_n']}", str(lead["lead_n"])]),
                "Created On": mess_date(rng, lead["created"], allow_blank=True),
                "Full Name": name,
                "Src": src,
                "Lead Status": status,
                "Est. Value": mess_money(rng, lead["est"], allow_blank=True),
                "Notes": lead["notes"],
            })

    # Disclaimer
    (RAW / "README.txt").write_text(
        "SAMPLE / SYNTHETIC DATA for a fictional business: Lantern Key Pool & Air "
        "(Delray Beach / Palm Beach County, FL area).\n"
        "No real customers, invoices, or companies. Emails use the .invalid TLD. "
        "Phone numbers use the 555 prefix reserved for fiction.\n"
        "Window: 2026-01-06 through 2026-11-20. As-of date for AR aging: 2026-11-25.\n"
        "Files:\n"
        "  invoices_payments.csv  — invoices and payments in one messy export\n"
        "  jobs_expenses.xlsx     — Jobs Export + Expenses sheets (headers not on row 1)\n"
        "  crm_leads.csv          — CRM-style lead export\n"
        "These files are intentionally messy: mixed date and currency formats, "
        "duplicate customer name variants, blank fields, duplicate rows, and a few bad values.\n",
        encoding="utf-8",
    )

    print(f"Wrote raw exports to {RAW}")
    print(f"  Truth invoices: {len(invoices)}")
    print(f"  Truth payments: {len(payments)}")
    print(f"  Truth jobs (completed, matched): {len(jobs)}")
    print(f"  Invoice CSV rows (incl. payment splits and planted dups): {len(inv_rows)}")
    print(f"  Expenses: {len(expenses)} (+ possible dup in Excel)")
    print(f"  Leads: {len(leads)}")
    print(f"  Duplicate invoice rows planted for ids: {sorted(dup_targets)}")
    print(f"  Unparseable amounts: {sorted(bad_amount_ids)}")
    print(f"  Missing/bad dates: {sorted(bad_date_ids)}")
    print(f"  Typo year 2027 invoice: {typo_year_id}")
    paid = sum(inv["amount"] for inv in invoices if inv["status_plan"] == "Paid")
    open_amt = sum(inv["amount"] for inv in invoices if inv["status_plan"] == "Open")
    partial = [inv for inv in invoices if inv["status_plan"] == "Partial"]
    print(f"  Status Paid billed: {paid:,.2f}  Open billed: {open_amt:,.2f}  Partials: {len(partial)}")


if __name__ == "__main__":
    main()
