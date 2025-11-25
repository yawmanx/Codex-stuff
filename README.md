# Codex-stuff

This repository contains test assets for Codex interactive pathways.

## Enhanced CV

An enhanced Markdown edition of **Yaw Manso Kyeremeh's** curriculum vitae is available in [`Yaw_Manso_Kyeremeh_CV.md`](Yaw_Manso_Kyeremeh_CV.md). It restructures the original PDF content with clearer sectioning, quantified impact, and modernised phrasing for easier reuse in web and document workflows.

## Church of Pentecost finance and membership CLI

The `church_finance_manager.py` script provides a lightweight, SQLite-backed way to track members, attendance, income (tithes, offerings, missions, ministry departments), expenses, pledges, and pledge redemptions for the Church of Pentecost.

### Quick start

1. **Create the database:**

   ```bash
   python church_finance_manager.py init-db
   ```

2. **Add members and attendance:**

   ```bash
   python church_finance_manager.py add-member "Mary Mensah" --contact "+233-555-000000"
   python church_finance_manager.py record-attendance 1 2024-06-09 --notes "Sunday service"
   ```

3. **Capture income and expenses:**

   ```bash
   python church_finance_manager.py add-income 2024-06-09 "Tithe" 250.00 --source "Second service"
   python church_finance_manager.py add-income 2024-06-09 "Missions" 120.00 --source "Missions offering"
   python church_finance_manager.py add-expense 2024-06-10 "Groceries" 80.00 --payee "Communion supplies"
   ```

4. **Track pledges and redemptions:**

   ```bash
   python church_finance_manager.py add-pledge 1 500.00 "Building Fund" 2024-06-09 --due 2024-09-30
   python church_finance_manager.py pay-pledge 1 150.00 2024-06-23 --notes "Part payment"
   ```

5. **Print reports (and optionally save to a file):**

   ```bash
   python church_finance_manager.py report --start 2024-06-01 --end 2024-06-30 --output june-report.txt
   ```

Use `--db path/to/file.db` with any command to store data in a different location.
