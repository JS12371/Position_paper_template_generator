import sqlite3
from datetime import datetime

DB_NAME = "database.db"

def init_db():
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.executescript("""
    CREATE TABLE IF NOT EXISTS Cases (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        stratifier TEXT NOT NULL,
        created_at TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS Exhibits (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        title TEXT UNIQUE NOT NULL
    );

    CREATE TABLE IF NOT EXISTS CaseExhibits (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        case_id INTEGER NOT NULL,
        exhibit_id INTEGER NOT NULL,
        FOREIGN KEY(case_id) REFERENCES Cases(id),
        FOREIGN KEY(exhibit_id) REFERENCES Exhibits(id)
    );
    """)
    conn.commit()
    conn.close()


def add_case(stratifier, created_at, exhibit_titles):
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()

    # Insert the case
    c.execute("INSERT INTO Cases (stratifier, created_at) VALUES (?, ?)", (stratifier, created_at))
    case_id = c.lastrowid

    # Handle exhibits
    for title in exhibit_titles:
        c.execute("INSERT OR IGNORE INTO Exhibits (title) VALUES (?)", (title,))
        c.execute("SELECT id FROM Exhibits WHERE title = ?", (title,))
        exhibit_id = c.fetchone()[0]
        c.execute("INSERT INTO CaseExhibits (case_id, exhibit_id) VALUES (?, ?)", (case_id, exhibit_id))

    conn.commit()
    conn.close()
    print(f"✅ Added case under {stratifier} with {len(exhibit_titles)} exhibit(s).")


def search_exhibits(mode, stratifier, from_date=None):
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()

    # Select matching cases
    if mode == "ALL_AFTER_DATE":
        c.execute("SELECT id FROM Cases WHERE stratifier = ? AND created_at > ?", (stratifier, from_date))
    else:
        c.execute("SELECT id FROM Cases WHERE stratifier = ?", (stratifier,))
    case_ids = [row[0] for row in c.fetchall()]
    
    if not case_ids:
        print("⚠️ No cases found.")
        return

    # Get exhibit sets
    exhibit_sets = []
    for cid in case_ids:
        c.execute("""
            SELECT E.title FROM CaseExhibits CE
            JOIN Exhibits E ON E.id = CE.exhibit_id
            WHERE CE.case_id = ?
        """, (cid,))
        exhibits = set([row[0] for row in c.fetchall()])
        exhibit_sets.append(exhibits)

    if mode.startswith("ALL"):
        result = set.intersection(*exhibit_sets)
    elif mode == "ANY":
        result = set.union(*exhibit_sets)

    print(f"\n📄 {mode} - {stratifier}" + (f" after {from_date}" if from_date else ""))
    for r in sorted(result):
        print(f"  - {r}")
    print(f"\n🧾 {len(result)} exhibit(s) found.\n")

    conn.close()


def main():
    init_db()
    while True:
        print("\n📘 Exhibit DB - Menu")
        print("1. Add new case")
        print("2. Search exhibits")
        print("3. Exit")
        choice = input("Choose option: ")

        if choice == "1":
            stratifier = input("Enter stratifier type: ").strip()
            created_at = input("Enter date (YYYY-MM-DD): ").strip()
            print("Enter exhibit titles one per line. Type 'STOP' to finish.")
            exhibits = []
            while True:
                exhibit = input("Exhibit: ").strip()
                if exhibit.upper() == "STOP":
                    break
                if exhibit:
                    exhibits.append(exhibit)

            add_case(stratifier, created_at, exhibits)

        elif choice == "2":
            print("\nSearch Modes: ALL (all cases with this stratifier have these exhibits), ANY (there exists a case with this stratifier that has at least one of these exhibits), ALL_AFTER_DATE (all cases with this stratifier after this date have these exhibits)")
            mode = input("Enter search mode: ").strip().upper()
            stratifier = input("Enter stratifier: ").strip()
            from_date = None
            if mode == "ALL_AFTER_DATE":
                from_date = input("Enter start date (YYYY-MM-DD): ").strip()
            search_exhibits(mode, stratifier, from_date)

        elif choice == "3":
            print("👋 Exiting...")
            break
        else:
            print("❌ Invalid option")


if __name__ == "__main__":
    main()
