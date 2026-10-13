#!/usr/bin/env python3
"""Generate a synthetic user-access export with deliberately planted problems.

Everything here is fake. Besides data/sample_users.csv it writes data/answer_key.csv
listing which account was planted for which rule, so you can prove the tool catches them.

Usage:  python make_sample_data.py [--as-of YYYY-MM-DD] [--seed 7]
"""
import argparse
import csv
import random
from datetime import date, timedelta
from pathlib import Path

FIRST = ["Alex", "Priya", "Marcus", "Elena", "Jordan", "Wei", "Fatima", "Dmitri", "Grace", "Omar",
         "Sofia", "Liam", "Aisha", "Noah", "Mei", "Carlos", "Hannah", "Ravi", "Tessa", "Ben"]
LAST = ["Carter", "Nguyen", "Okafor", "Brooks", "Alvarez", "Kim", "Fischer", "Patel", "Reyes", "Khan",
        "Murphy", "Tanaka", "Rossi", "Haddad", "Larsen", "Singh", "Duarte", "Evans", "Zhou", "Moreau"]

DEPTS = {
    "Finance": ["Accountant", "Financial Analyst", "AP Specialist", "Controller"],
    "IT": ["Systems Admin", "Help Desk Analyst", "Network Engineer", "DevOps Engineer"],
    "HR": ["HR Generalist", "Recruiter", "Benefits Analyst"],
    "Sales": ["Account Executive", "Sales Engineer", "SDR"],
    "Operations": ["Ops Analyst", "Logistics Coordinator", "Project Manager"],
    "Legal": ["Paralegal", "Counsel", "Compliance Analyst"],
}
BASE_GROUPS = {
    "Finance": ["Finance-Users"], "IT": ["IT-Users"], "HR": ["HR-Users"],
    "Sales": ["Sales-Users"], "Operations": ["Ops-Users"], "Legal": ["Legal-Users"],
}
MANAGERS = {"Finance": "d.kim", "IT": "p.nguyen", "HR": "s.okafor",
            "Sales": "t.brooks", "Operations": "m.alvarez", "Legal": "l.fischer"}

COLUMNS = ["username", "display_name", "department", "job_title", "manager", "account_status",
           "last_login", "created_date", "groups", "privileged", "privileged_justified_date",
           "account_type", "termination_date", "expiry_date", "interactive_login"]

TOTAL_USERS = 200


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--as-of", default=date.today().isoformat())
    ap.add_argument("--seed", type=int, default=7)
    ap.add_argument("--out", default="data")
    args = ap.parse_args()

    as_of = date.fromisoformat(args.as_of)
    rng = random.Random(args.seed)
    used, rows, key = set(), [], []

    def ago(days):
        return (as_of - timedelta(days=days)).isoformat()

    def ahead(days):
        return (as_of + timedelta(days=days)).isoformat()

    def unique(name):
        cand, n = name, 1
        while cand in used:
            n += 1
            cand = f"{name}{n}"
        used.add(cand)
        return cand

    def normal(dept=None):
        dept = dept or rng.choice(list(DEPTS))
        first, last = rng.choice(FIRST), rng.choice(LAST)
        groups = list(BASE_GROUPS[dept])
        if dept == "Finance":                      # never both Submitter and Approver
            roll = rng.random()
            if roll < 0.2:
                groups.append("Payments Submitter")
            elif roll < 0.3:
                groups.append("Payments Approver")
        return {
            "username": unique(first[0].lower() + last.lower()),
            "display_name": f"{first} {last}", "department": dept,
            "job_title": rng.choice(DEPTS[dept]), "manager": MANAGERS[dept],
            "account_status": "Enabled", "last_login": ago(rng.randint(0, 30)),
            "created_date": ago(rng.randint(365, 365 * 8)), "groups": ";".join(groups),
            "privileged": "no", "privileged_justified_date": "", "account_type": "employee",
            "termination_date": "", "expiry_date": "", "interactive_login": "yes",
        }

    def add(row, rule=None):
        rows.append(row)
        if rule:
            key.append((row["username"], rule))

    # R01 inactive but enabled (2 have never logged in)
    for i in range(12):
        r = normal()
        r["last_login"] = "" if i < 2 else ago(rng.randint(100, 400))
        add(r, "R01")

    # R02 terminated but still enabled
    for _ in range(6):
        r, days = normal(), rng.randint(5, 120)
        r["termination_date"], r["last_login"] = ago(days), ago(days + rng.randint(0, 5))
        add(r, "R02")
    # clean: terminated and correctly disabled
    for _ in range(8):
        r, days = normal(), rng.randint(5, 300)
        r.update(account_status="Disabled", termination_date=ago(days), last_login=ago(days))
        add(r)

    # R03 contractors past expiry (2 with no expiry set at all) + valid contractors
    for i in range(10):
        r = normal()
        r.update(account_type="contractor", job_title="Contractor - " + r["job_title"],
                 created_date=ago(rng.randint(200, 700)),
                 expiry_date="" if i < 2 else ago(rng.randint(10, 90)))
        add(r, "R03")
    for _ in range(10):
        r = normal()
        r.update(account_type="contractor", job_title="Contractor - " + r["job_title"],
                 created_date=ago(rng.randint(30, 300)), expiry_date=ahead(rng.randint(15, 180)))
        add(r)

    # R04 privileged without recent justification (half blank, half stale) + justified admins
    for i in range(8):
        r = normal("IT")
        r.update(privileged="yes", groups="IT-Users;Domain Admins",
                 privileged_justified_date="" if i < 4 else ago(rng.randint(250, 700)))
        add(r, "R04")
    for _ in range(8):
        r = normal("IT")
        r.update(privileged="yes", groups="IT-Users;Server Admins",
                 privileged_justified_date=ago(rng.randint(10, 150)))
        add(r)

    # R05 segregation-of-duties conflicts
    for _ in range(6):
        r = normal("Finance")
        r["groups"] = "Finance-Users;Payments Submitter;Payments Approver"
        add(r, "R05")

    # R06 no manager
    for _ in range(7):
        r = normal()
        r["manager"] = ""
        add(r, "R06")

    # R07 service accounts with interactive login + clean service accounts
    svc = ["backup", "sqlreport", "etl", "monitor", "printsvc", "webapp", "scanner-sync", "hrfeed",
           "ldapbind", "patching", "vpnauth", "crm-sync", "mailrelay", "siem", "build", "dbadmin"]
    for i, name in enumerate(svc):
        r = normal("IT")
        r.update(username=unique("svc-" + name), display_name=f"Service {name}",
                 job_title="Service account", groups="Service-Accounts", account_type="service",
                 manager=MANAGERS["IT"], created_date=ago(rng.randint(300, 2000)),
                 interactive_login="yes" if i < 6 else "no")
        add(r, "R07" if i < 6 else None)

    # R08 shared / generic accounts
    for name in ["shared.helpdesk", "generic.scanner", "admin", "guest.wifi"]:
        r = normal("IT")
        r.update(username=unique(name), display_name=name.replace(".", " ").title(),
                 job_title="Shared account", account_type="shared", manager=MANAGERS["IT"])
        add(r, "R08")

    while len(rows) < TOTAL_USERS:
        add(normal())

    rng.shuffle(rows)
    out = Path(args.out)
    out.mkdir(parents=True, exist_ok=True)
    with open(out / "sample_users.csv", "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=COLUMNS)
        w.writeheader()
        w.writerows(rows)
    with open(out / "answer_key.csv", "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["username", "planted_rule"])
        w.writerows(key)
    print(f"Wrote {len(rows)} users and {len(key)} planted issues to {out}/ (as-of {as_of})")


if __name__ == "__main__":
    main()
