#!/usr/bin/env python3
"""
Relay: scoring and Excel builder.

Stages:
  prescore        Score people on title, seniority, department (max 85).
                  Writes scores back into people.json and prints the top N.
  final           Add email points (max 15), compute Relevance %, write the .xlsx.
  read-approvals  Read the reviewed .xlsx and print the rows marked Approve? = Y as JSON.

Pass --suppression relay_data/suppression.json to skip people who opted out or bounced.
Missing seniority or department data gets half credit until enrichment fills it in.

Scoring (out of 100):
  Title match       40  (title_match 0 to 1, set by Claude)
  Seniority fit     25  (exact level = full, one level off = half)
  Department match  20  (any overlap with persona departments)
  Email             15  (verified AND domain match = 15, one of the two = 7)

Requires: pip install openpyxl
"""

import argparse
import json
import sys
from datetime import datetime
from pathlib import Path

SENIORITY_ORDER = [
    "intern", "entry", "senior", "manager", "director",
    "head", "vp", "partner", "c_suite", "founder", "owner",
]

W_TITLE, W_SENIORITY, W_DEPT, W_EMAIL = 40, 25, 20, 15


def load_json(path):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def save_json(path, data):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)


def norm(s):
    return (s or "").strip().lower()


def norm_dept(d):
    d = norm(d)
    if d.startswith("master_"):
        d = d[len("master_"):]
    return d.replace("-", "_").replace(" ", "_")


def is_excluded(title, exclude_terms):
    t = norm(title)
    return any(norm(term) and norm(term) in t for term in exclude_terms)


UNKNOWN_CREDIT = 0.5


def load_suppression(path):
    """Return lookup sets from suppression.json. Missing file means an empty list."""
    sup = {"emails": set(), "ids": set(), "name_domain": set()}
    if not path or not Path(path).exists():
        return sup
    for e in load_json(path) or []:
        if e.get("email"):
            sup["emails"].add(norm(e["email"]))
        if e.get("apollo_id"):
            sup["ids"].add(str(e["apollo_id"]))
        if e.get("name") and e.get("company_domain"):
            sup["name_domain"].add((norm(e["name"]), norm(e["company_domain"]).removeprefix("www.")))
    return sup


def is_suppressed(person, sup, target_domain):
    if str(person.get("apollo_id") or "") in sup["ids"]:
        return True
    if person.get("email") and norm(person["email"]) in sup["emails"]:
        return True
    key = (norm(person.get("name")), norm(target_domain).removeprefix("www."))
    return key in sup["name_domain"]


def seniority_points(person_level, target_levels):
    p = norm(person_level)
    targets = [norm(x) for x in target_levels]
    if not targets:
        return 0.0
    if not p:
        return UNKNOWN_CREDIT
    if p in targets:
        return 1.0
    if p in SENIORITY_ORDER:
        pi = SENIORITY_ORDER.index(p)
        for t in targets:
            if t in SENIORITY_ORDER and abs(SENIORITY_ORDER.index(t) - pi) == 1:
                return 0.5
    return 0.0


def dept_points(person_depts, target_depts):
    p = {norm_dept(d) for d in (person_depts or []) if d}
    t = {norm_dept(d) for d in (target_depts or []) if d}
    if not t:
        return 0.0
    if not p:
        return UNKNOWN_CREDIT
    if p & t:
        return 1.0
    for a in p:
        for b in t:
            if a in b or b in a:
                return 1.0
    return 0.0


def email_domain_matches(email, target_domain):
    if not email or "@" not in email:
        return False
    dom = norm(email.split("@", 1)[1])
    target = norm(target_domain).removeprefix("www.")
    return dom == target or dom.endswith("." + target)


def clamp01(x):
    try:
        x = float(x)
    except (TypeError, ValueError):
        return 0.0
    return max(0.0, min(1.0, x))


def prescore(people, persona, sup, target_domain):
    excludes = persona.get("exclude_titles", [])
    for p in people:
        p["suppressed"] = is_suppressed(p, sup, target_domain)
        if p["suppressed"] or is_excluded(p.get("title"), excludes):
            p["excluded"] = True
            p["prescore"] = 0.0
            continue
        p["excluded"] = False
        score = (
            clamp01(p.get("title_match")) * W_TITLE
            + seniority_points(p.get("seniority"), persona.get("seniorities", [])) * W_SENIORITY
            + dept_points(p.get("departments"), persona.get("departments", [])) * W_DEPT
        )
        p["prescore"] = round(score, 1)
    people.sort(key=lambda x: x.get("prescore", 0), reverse=True)
    rank = 0
    for p in people:
        if not p["excluded"]:
            rank += 1
            p["rank"] = rank
        else:
            p["rank"] = None
    return people


def final_score(people, target_domain):
    for p in people:
        email = p.get("email") or ""
        verified = norm(p.get("email_status")) == "verified"
        domain_ok = email_domain_matches(email, target_domain)
        p["verified"] = verified
        p["domain_match"] = domain_ok
        p["usable_email"] = bool(email) and verified and domain_ok
        if not email:
            pts = 0
        elif verified and domain_ok:
            pts = W_EMAIL
        elif verified or domain_ok:
            pts = 7
        else:
            pts = 0
        base = p.get("prescore")
        if base is None:
            base = 0.0
        p["relevance"] = 0 if p.get("excluded") else int(round(base + pts))
        if p.get("suppressed"):
            p["usable_email"] = False
            p["notes"] = "On suppression list"
        elif p.get("excluded"):
            p["notes"] = "Excluded title"
        elif not email:
            p["notes"] = "No email revealed"
        elif not verified:
            p["notes"] = "Email not verified"
        elif not domain_ok:
            p["notes"] = "Email domain does not match target"
        else:
            p["notes"] = "Ready"
    people.sort(key=lambda x: x.get("relevance", 0), reverse=True)
    return people


def write_xlsx(people, out_path, persona, seller, target, target_domain):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.worksheet.datavalidation import DataValidation

    wb = Workbook()
    ws = wb.active
    ws.title = "Contacts"

    headers = ["Name", "Title", "Seniority", "Department", "Email", "Verified",
               "Domain Match", "Relevance %", "LinkedIn", "Approve?", "Notes"]
    ws.append(headers)

    header_fill = PatternFill("solid", fgColor="1F2937")
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
        cell.alignment = Alignment(vertical="center")

    green = PatternFill("solid", fgColor="C6EFCE")
    yellow = PatternFill("solid", fgColor="FFEB9C")
    red = PatternFill("solid", fgColor="FFC7CE")

    rows = [p for p in people if not p.get("excluded")]
    for p in rows:
        depts = ", ".join(norm_dept(d).replace("_", " ") for d in (p.get("departments") or []))
        ws.append([
            p.get("name", ""),
            p.get("title", ""),
            (p.get("seniority") or "").replace("_", " "),
            depts,
            p.get("email") or "",
            "Yes" if p.get("verified") else "No",
            "Yes" if p.get("domain_match") else "No",
            p.get("relevance", 0),
            p.get("linkedin_url") or "",
            "",
            p.get("notes", ""),
        ])
        r = ws.max_row
        score_cell = ws.cell(row=r, column=8)
        s = p.get("relevance", 0)
        score_cell.fill = green if s >= 75 else yellow if s >= 50 else red
        url = p.get("linkedin_url")
        if url:
            link = ws.cell(row=r, column=9)
            link.hyperlink = url
            link.font = Font(color="0563C1", underline="single")

    if rows:
        dv = DataValidation(type="list", formula1='"Y,N"', allow_blank=True)
        ws.add_data_validation(dv)
        dv.add(f"J2:J{ws.max_row}")

    widths = [24, 34, 12, 22, 34, 10, 14, 13, 40, 11, 32]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = w
    ws.freeze_panes = "A2"

    s = wb.create_sheet("Run summary")
    seller = seller or {}
    target = target or {}
    summary = [
        ("Generated", datetime.now().strftime("%Y-%m-%d %H:%M")),
        ("", ""),
        ("Seller", seller.get("company", "")),
        ("Seller domain", seller.get("domain", "")),
        ("What they sell", seller.get("offering", "")),
        ("Problem solved", seller.get("problem_solved", "")),
        ("Buyer function", seller.get("buyer_function", "")),
        ("", ""),
        ("Target", target.get("name", "")),
        ("Target domain", target_domain),
        ("Industry", target.get("industry", "")),
        ("Employees", target.get("employee_count", "")),
        ("", ""),
        ("Titles targeted", ", ".join(persona.get("titles", []))),
        ("Seniority", ", ".join(persona.get("seniorities", []))),
        ("Departments", ", ".join(persona.get("departments", []))),
        ("Excluded", ", ".join(persona.get("exclude_titles", []))),
        ("Persona reasoning", persona.get("reasoning", "")),
        ("", ""),
        ("People in sheet", len(rows)),
        ("Usable emails", sum(1 for p in rows if p.get("usable_email"))),
        ("Excluded by title", sum(1 for p in people if p.get("excluded") and not p.get("suppressed"))),
        ("Suppressed (opted out or bounced)", sum(1 for p in people if p.get("suppressed"))),
        ("", ""),
        ("Scoring", "Title 40, Seniority 25, Department 20, Email 15"),
    ]
    for k, v in summary:
        s.append([k, v])
        if k:
            s.cell(row=s.max_row, column=1).font = Font(bold=True)
    s.column_dimensions["A"].width = 20
    s.column_dimensions["B"].width = 70
    for row in s.iter_rows(min_col=2, max_col=2):
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")

    Path(out_path).parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)


def read_approvals(xlsx_path):
    from openpyxl import load_workbook
    wb = load_workbook(xlsx_path, data_only=True)
    ws = wb["Contacts"]
    headers = [c.value for c in ws[1]]
    out = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        rec = dict(zip(headers, row))
        if norm(str(rec.get("Approve?") or "")) in ("y", "yes"):
            out.append({
                "name": rec.get("Name"),
                "title": rec.get("Title"),
                "email": rec.get("Email"),
                "relevance": rec.get("Relevance %"),
                "usable_email": rec.get("Notes") == "Ready",
                "notes": rec.get("Notes"),
            })
    return out


def print_table(people, n, final=False):
    shown = [p for p in people if not p.get("excluded")][:n]
    key = "relevance" if final else "prescore"
    label = "Relevance %" if final else "Pre-score"
    print(f"\nTop {len(shown)} by {label}:")
    print(f"{'#':>3}  {label:>11}  {'Name':<24} {'Title':<36} apollo_id")
    for i, p in enumerate(shown, start=1):
        print(f"{i:>3}  {p.get(key, 0):>11}  {str(p.get('name',''))[:24]:<24} "
              f"{str(p.get('title',''))[:36]:<36} {p.get('apollo_id','')}")


def main():
    ap = argparse.ArgumentParser(description="Relay scoring and Excel builder")
    ap.add_argument("people", help="Path to people.json (or the .xlsx for read-approvals)")
    ap.add_argument("--stage", required=True, choices=["prescore", "final", "read-approvals"])
    ap.add_argument("--persona", help="Path to persona.json")
    ap.add_argument("--target-domain", help="Target company domain, e.g. b.com")
    ap.add_argument("--seller", help="Path to seller_profile.json")
    ap.add_argument("--target", help="Path to target_profile.json")
    ap.add_argument("--out", help="Output .xlsx path (final stage)")
    ap.add_argument("--suppression", help="Path to suppression.json (people never to contact)")
    ap.add_argument("--top-n", type=int, default=10)
    args = ap.parse_args()

    if args.stage == "read-approvals":
        print(json.dumps(read_approvals(args.people), indent=2))
        return

    if not args.persona or not args.target_domain:
        sys.exit("--persona and --target-domain are required for prescore and final")

    people = load_json(args.people)
    persona = load_json(args.persona)
    sup = load_suppression(args.suppression)

    if args.stage == "prescore":
        people = prescore(people, persona, sup, args.target_domain)
        save_json(args.people, people)
        print_table(people, args.top_n)
        n_sup = sum(1 for p in people if p.get("suppressed"))
        if n_sup:
            print(f"\nSuppressed (skipped, never enriched): {n_sup}")
        top_ids = [p.get("apollo_id") for p in people if not p.get("excluded")][: args.top_n]
        print("\nTOP_N_APOLLO_IDS=" + json.dumps(top_ids))
        return

    if not args.out:
        sys.exit("--out is required for the final stage")
    # Recompute with enriched data (seniority, departments, emails may have arrived).
    people = prescore(people, persona, sup, args.target_domain)
    seller = load_json(args.seller) if args.seller else {}
    target = load_json(args.target) if args.target else {}
    people = final_score(people, args.target_domain)
    save_json(args.people, people)
    write_xlsx(people, args.out, persona, seller, target, args.target_domain)
    print_table(people, 5, final=True)
    usable = sum(1 for p in people if p.get("usable_email"))
    n_sup = sum(1 for p in people if p.get("suppressed"))
    print(f"\nWrote {args.out}  |  people: {sum(1 for p in people if not p.get('excluded'))}"
          f"  |  usable emails: {usable}  |  suppressed: {n_sup}")


if __name__ == "__main__":
    main()
