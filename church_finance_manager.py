"""
Church of Pentecost Finance and Membership Manager
-------------------------------------------------
A lightweight CLI tool for tracking income, expenses, members, attendance,
pledges, and pledge redemptions for a church.

Data is stored in a local SQLite database (default: church_finance.db). Run
"python church_finance_manager.py init-db" once to create the schema.
"""
from __future__ import annotations

import argparse
import sqlite3
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Iterable, Optional

DB_NAME = "church_finance.db"


@dataclass
class IncomeSummary:
    category: str
    total: float


@dataclass
class ExpenseSummary:
    category: str
    total: float


@dataclass
class PledgeStatus:
    id: int
    member_name: str
    purpose: str
    pledged_amount: float
    redeemed_amount: float
    outstanding: float


def get_connection(db_path: Path) -> sqlite3.Connection:
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    return conn


def init_db(conn: sqlite3.Connection) -> None:
    cur = conn.cursor()
    cur.executescript(
        """
        PRAGMA foreign_keys = ON;

        CREATE TABLE IF NOT EXISTS members (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            contact TEXT,
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
        );

        CREATE TABLE IF NOT EXISTS attendance (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            member_id INTEGER NOT NULL,
            attended_on TEXT NOT NULL,
            notes TEXT,
            FOREIGN KEY(member_id) REFERENCES members(id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS incomes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            received_on TEXT NOT NULL,
            category TEXT NOT NULL,
            amount REAL NOT NULL,
            source TEXT,
            notes TEXT
        );

        CREATE TABLE IF NOT EXISTS expenses (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            spent_on TEXT NOT NULL,
            category TEXT NOT NULL,
            amount REAL NOT NULL,
            payee TEXT,
            notes TEXT
        );

        CREATE TABLE IF NOT EXISTS pledges (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            member_id INTEGER NOT NULL,
            purpose TEXT NOT NULL,
            amount REAL NOT NULL,
            pledged_on TEXT NOT NULL,
            due_on TEXT,
            notes TEXT,
            FOREIGN KEY(member_id) REFERENCES members(id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS pledge_payments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            pledge_id INTEGER NOT NULL,
            paid_on TEXT NOT NULL,
            amount REAL NOT NULL,
            notes TEXT,
            FOREIGN KEY(pledge_id) REFERENCES pledges(id) ON DELETE CASCADE
        );
        """
    )
    conn.commit()


def add_member(conn: sqlite3.Connection, name: str, contact: Optional[str]) -> int:
    cur = conn.cursor()
    cur.execute("INSERT INTO members (name, contact) VALUES (?, ?)", (name, contact))
    conn.commit()
    return cur.lastrowid


def list_members(conn: sqlite3.Connection) -> list[sqlite3.Row]:
    cur = conn.cursor()
    cur.execute("SELECT id, name, contact, created_at FROM members ORDER BY name")
    return cur.fetchall()


def record_attendance(conn: sqlite3.Connection, member_id: int, attended_on: str, notes: Optional[str]) -> None:
    _validate_date(attended_on)
    cur = conn.cursor()
    cur.execute(
        "INSERT INTO attendance (member_id, attended_on, notes) VALUES (?, ?, ?)",
        (member_id, attended_on, notes),
    )
    conn.commit()


def add_income(
    conn: sqlite3.Connection,
    received_on: str,
    category: str,
    amount: float,
    source: Optional[str],
    notes: Optional[str],
) -> None:
    _validate_date(received_on)
    cur = conn.cursor()
    cur.execute(
        "INSERT INTO incomes (received_on, category, amount, source, notes) VALUES (?, ?, ?, ?, ?)",
        (received_on, category, amount, source, notes),
    )
    conn.commit()


def add_expense(
    conn: sqlite3.Connection,
    spent_on: str,
    category: str,
    amount: float,
    payee: Optional[str],
    notes: Optional[str],
) -> None:
    _validate_date(spent_on)
    cur = conn.cursor()
    cur.execute(
        "INSERT INTO expenses (spent_on, category, amount, payee, notes) VALUES (?, ?, ?, ?, ?)",
        (spent_on, category, amount, payee, notes),
    )
    conn.commit()


def add_pledge(
    conn: sqlite3.Connection,
    member_id: int,
    purpose: str,
    amount: float,
    pledged_on: str,
    due_on: Optional[str],
    notes: Optional[str],
) -> int:
    _validate_date(pledged_on)
    if due_on:
        _validate_date(due_on)
    cur = conn.cursor()
    cur.execute(
        """
        INSERT INTO pledges (member_id, purpose, amount, pledged_on, due_on, notes)
        VALUES (?, ?, ?, ?, ?, ?)
        """,
        (member_id, purpose, amount, pledged_on, due_on, notes),
    )
    conn.commit()
    return cur.lastrowid


def add_pledge_payment(
    conn: sqlite3.Connection,
    pledge_id: int,
    paid_on: str,
    amount: float,
    notes: Optional[str],
) -> None:
    _validate_date(paid_on)
    cur = conn.cursor()
    cur.execute(
        "INSERT INTO pledge_payments (pledge_id, paid_on, amount, notes) VALUES (?, ?, ?, ?)",
        (pledge_id, paid_on, amount, notes),
    )
    conn.commit()


def summarize_income(conn: sqlite3.Connection, start: Optional[str], end: Optional[str]) -> Iterable[IncomeSummary]:
    cur = conn.cursor()
    query = "SELECT category, SUM(amount) as total FROM incomes"
    params: list[str] = []
    where = []
    if start:
        _validate_date(start)
        where.append("received_on >= ?")
        params.append(start)
    if end:
        _validate_date(end)
        where.append("received_on <= ?")
        params.append(end)
    if where:
        query += " WHERE " + " AND ".join(where)
    query += " GROUP BY category ORDER BY category"
    cur.execute(query, params)
    for row in cur.fetchall():
        yield IncomeSummary(category=row[0], total=row[1] or 0.0)


def summarize_expenses(conn: sqlite3.Connection, start: Optional[str], end: Optional[str]) -> Iterable[ExpenseSummary]:
    cur = conn.cursor()
    query = "SELECT category, SUM(amount) as total FROM expenses"
    params: list[str] = []
    where = []
    if start:
        _validate_date(start)
        where.append("spent_on >= ?")
        params.append(start)
    if end:
        _validate_date(end)
        where.append("spent_on <= ?")
        params.append(end)
    if where:
        query += " WHERE " + " AND ".join(where)
    query += " GROUP BY category ORDER BY category"
    cur.execute(query, params)
    for row in cur.fetchall():
        yield ExpenseSummary(category=row[0], total=row[1] or 0.0)


def pledge_statuses(conn: sqlite3.Connection) -> Iterable[PledgeStatus]:
    cur = conn.cursor()
    cur.execute(
        """
        SELECT
            p.id,
            m.name AS member_name,
            p.purpose,
            p.amount AS pledged_amount,
            IFNULL(SUM(pp.amount), 0) AS redeemed_amount,
            p.amount - IFNULL(SUM(pp.amount), 0) AS outstanding
        FROM pledges p
        JOIN members m ON p.member_id = m.id
        LEFT JOIN pledge_payments pp ON p.id = pp.pledge_id
        GROUP BY p.id, m.name, p.purpose, p.amount
        ORDER BY outstanding DESC
        """
    )
    for row in cur.fetchall():
        yield PledgeStatus(
            id=row[0],
            member_name=row[1],
            purpose=row[2],
            pledged_amount=row[3],
            redeemed_amount=row[4],
            outstanding=row[5],
        )


def print_report(
    conn: sqlite3.Connection,
    start: Optional[str],
    end: Optional[str],
    output: Optional[Path],
) -> None:
    lines = [
        "Church of Pentecost — Finance & Attendance Summary",
        f"Period: {start or 'beginning'} to {end or 'latest'}",
        "",
        "Income by Category:",
    ]
    total_income = 0.0
    for row in summarize_income(conn, start, end):
        lines.append(f"  - {row.category}: GH₵{row.total:,.2f}")
        total_income += row.total
    lines.append(f"  Total income: GH₵{total_income:,.2f}\n")

    lines.append("Expenses by Category:")
    total_expense = 0.0
    for row in summarize_expenses(conn, start, end):
        lines.append(f"  - {row.category}: GH₵{row.total:,.2f}")
        total_expense += row.total
    lines.append(f"  Total expenses: GH₵{total_expense:,.2f}\n")

    lines.append("Pledge Status:")
    for pledge in pledge_statuses(conn):
        lines.append(
            "  - "
            f"Pledge #{pledge.id} by {pledge.member_name} for {pledge.purpose}: "
            f"pledged GH₵{pledge.pledged_amount:,.2f}, "
            f"redeemed GH₵{pledge.redeemed_amount:,.2f}, "
            f"outstanding GH₵{pledge.outstanding:,.2f}"
        )

    report = "\n".join(lines)
    print(report)

    if output:
        output.write_text(report, encoding="utf-8")
        print(f"\nReport saved to {output}")


def _validate_date(date_text: str) -> None:
    try:
        datetime.strptime(date_text, "%Y-%m-%d")
    except ValueError as exc:
        raise ValueError("Dates must use YYYY-MM-DD format.") from exc


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Church finance and membership manager")
    parser.add_argument("--db", type=Path, default=Path(DB_NAME), help="Path to SQLite database file")

    sub = parser.add_subparsers(dest="command", required=True)

    sub.add_parser("init-db", help="Create database tables")

    add_member_cmd = sub.add_parser("add-member", help="Register a new member")
    add_member_cmd.add_argument("name")
    add_member_cmd.add_argument("--contact")

    list_members_cmd = sub.add_parser("list-members", help="Show all members")

    attendance_cmd = sub.add_parser("record-attendance", help="Record attendance for a member")
    attendance_cmd.add_argument("member_id", type=int)
    attendance_cmd.add_argument("date", help="YYYY-MM-DD")
    attendance_cmd.add_argument("--notes")

    income_cmd = sub.add_parser("add-income", help="Record church income (tithes, offerings, missions, etc.)")
    income_cmd.add_argument("date", help="YYYY-MM-DD")
    income_cmd.add_argument("category", help="e.g., Tithe, Offering, Missions, Youth, Men, Women, Building Fund")
    income_cmd.add_argument("amount", type=float)
    income_cmd.add_argument("--source", help="Who gave or the service name")
    income_cmd.add_argument("--notes")

    expense_cmd = sub.add_parser("add-expense", help="Record a church expense")
    expense_cmd.add_argument("date", help="YYYY-MM-DD")
    expense_cmd.add_argument("category", help="e.g., Utilities, Groceries, Transport, Gifts, Maintenance")
    expense_cmd.add_argument("amount", type=float)
    expense_cmd.add_argument("--payee", help="Vendor or beneficiary")
    expense_cmd.add_argument("--notes")

    pledge_cmd = sub.add_parser("add-pledge", help="Record a pledge made by a member")
    pledge_cmd.add_argument("member_id", type=int)
    pledge_cmd.add_argument("amount", type=float)
    pledge_cmd.add_argument("purpose", help="e.g., Building Fund, Missions, Anniversary")
    pledge_cmd.add_argument("date", help="YYYY-MM-DD for pledge date")
    pledge_cmd.add_argument("--due", help="Optional YYYY-MM-DD due date")
    pledge_cmd.add_argument("--notes")

    pledge_payment_cmd = sub.add_parser("pay-pledge", help="Record a pledge redemption")
    pledge_payment_cmd.add_argument("pledge_id", type=int)
    pledge_payment_cmd.add_argument("amount", type=float)
    pledge_payment_cmd.add_argument("date", help="YYYY-MM-DD")
    pledge_payment_cmd.add_argument("--notes")

    report_cmd = sub.add_parser("report", help="Print income/expense and pledge summary")
    report_cmd.add_argument("--start", help="YYYY-MM-DD start date", default=None)
    report_cmd.add_argument("--end", help="YYYY-MM-DD end date", default=None)
    report_cmd.add_argument("--output", type=Path, help="Optional path to save the report")

    return parser.parse_args()


def main() -> None:
    args = parse_args()
    conn = get_connection(args.db)

    if args.command == "init-db":
        init_db(conn)
        print(f"Database initialized at {args.db}")
        return

    if args.command == "add-member":
        member_id = add_member(conn, args.name, args.contact)
        print(f"Member added with ID {member_id}")
        return

    if args.command == "list-members":
        for m in list_members(conn):
            contact = f" ({m['contact']})" if m["contact"] else ""
            print(f"#{m['id']}: {m['name']}{contact} — added {m['created_at']}")
        return

    if args.command == "record-attendance":
        record_attendance(conn, args.member_id, args.date, args.notes)
        print(f"Attendance recorded for member #{args.member_id} on {args.date}")
        return

    if args.command == "add-income":
        add_income(conn, args.date, args.category, args.amount, args.source, args.notes)
        print(
            "Income recorded: "
            f"{args.category} GH₵{args.amount:,.2f} on {args.date}"
        )
        return

    if args.command == "add-expense":
        add_expense(conn, args.date, args.category, args.amount, args.payee, args.notes)
        print(
            "Expense recorded: "
            f"{args.category} GH₵{args.amount:,.2f} on {args.date}"
        )
        return

    if args.command == "add-pledge":
        pledge_id = add_pledge(conn, args.member_id, args.purpose, args.amount, args.date, args.due, args.notes)
        print(f"Pledge #{pledge_id} recorded for member #{args.member_id}")
        return

    if args.command == "pay-pledge":
        add_pledge_payment(conn, args.pledge_id, args.date, args.amount, args.notes)
        print(f"Pledge payment recorded for pledge #{args.pledge_id}")
        return

    if args.command == "report":
        print_report(conn, args.start, args.end, args.output)
        return


if __name__ == "__main__":
    main()
