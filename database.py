import sqlite3
import os
import re
from datetime import datetime

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'tameer_crm.db')

SCAFFOLDING_COMPONENTS = [
    # Ledgers
    (1, "Ledgers", "Ledger 0.6m", "Pcs.", 15.00),
    (2, "Ledgers", "Ledger 0.9m", "Pcs.", 18.00),
    (3, "Ledgers", "Ledger 1.0m", "Pcs.", 20.00),
    (4, "Ledgers", "Ledger 1.2m", "Pcs.", 22.00),
    (5, "Ledgers", "Ledger 1.3m", "Pcs.", 24.00),
    (6, "Ledgers", "Ledger 1.5m", "Pcs.", 25.00),
    (7, "Ledgers", "Ledger 1.6m", "Pcs.", 26.00),
    (8, "Ledgers", "1.8m Ledger", "Pcs.", 28.00),
    (9, "Ledgers", "Ledger 2.3 m", "Pcs.", 32.00),
    (10, "Ledgers", "Ledger 2.0 m", "Pcs.", 30.00),
    (11, "Ledgers", "2.5m Ledger", "Pcs.", 35.00),

    # Standards
    (12, "Standards", "1.0m Standard", "Pcs.", 9.50),
    (13, "Standards", "Standard 1.3m", "Pcs.", 18.00),
    (14, "Standards", "Standard 1.5m", "Pcs.", 25.00),
    (15, "Standards", "Standard 2m", "Pcs.", 38.00),
    (16, "Standards", "Standard 2.3m", "Pcs.", 45.00),
    (17, "Standards", "Standard 2.5m", "Pcs.", 52.00),
    (18, "Standards", "Standard 2.8m", "Pcs.", 58.00),
    (19, "Standards", "3.0m Vertical Standard", "Pcs.", 64.00),

    # Intermediate Transoms
    (20, "Intermediate Transoms", "Intermediate Transom 0.9m", "Pcs.", 18.00),
    (21, "Intermediate Transoms", "Intermediate Transom 1.0m", "Pcs.", 20.00),
    (22, "Intermediate Transoms", "Intermediate Transom 1.2m", "Pcs.", 22.00),
    (23, "Intermediate Transoms", "Intermediate Transom 1.3m", "Pcs.", 24.00),
    (24, "Intermediate Transoms", "Intermediate Transom 1.5m", "Pcs.", 26.00),
    (25, "Intermediate Transoms", "Intermediate Transom 1.6m", "Pcs.", 27.00),
    (26, "Intermediate Transoms", "1.8m Intermediate Transom", "Pcs.", 29.00),
    (27, "Intermediate Transoms", "Intermediate Transom 2.0m", "Pcs.", 32.00),
    (28, "Intermediate Transoms", "Intermediate Transom 2.3m", "Pcs.", 36.00),
    (29, "Intermediate Transoms", "2.5m Intermediate Transom", "Pcs.", 40.00),

    # Jacks & Bases
    (30, "Jacks & Bases", "Adj. Base Jack 500mm", "Pcs.", 19.00),
    (31, "Jacks & Bases", "Basejack 660mm", "Pcs.", 24.00),
    (32, "Jacks & Bases", "Universal Jack", "Pcs.", 28.00),
    (33, "Jacks & Bases", "Forkhead", "Pcs.", 22.00),

    # Alu. Beam- Light duty
    (34, "Alu Beams (Light Duty)", "Alu. Beam- Light duty - 150x 6.0m", "Pcs.", 180.00),
    (35, "Alu Beams (Light Duty)", "Alu. Beam Light duty -150x5.5m", "Pcs.", 165.00),
    (36, "Alu Beams (Light Duty)", "Alu. Beam Light duty - 150x5.0m", "Pcs.", 150.00),
    (37, "Alu Beams (Light Duty)", "Alu. Beam Light duty -150X 4.5m", "Pcs.", 135.00),
    (38, "Alu Beams (Light Duty)", "Alu. Beam Light duty -150X 4.0m", "Pcs.", 120.00),
    (39, "Alu Beams (Light Duty)", "Alu. Beam Light duty -150X3.5m", "Pcs.", 105.00),
    (40, "Alu Beams (Light Duty)", "Alu. Beam Light duty -150X3.0m", "Pcs.", 90.00),
    (41, "Alu Beams (Light Duty)", "Alu. Beam Light duty -150X2.5m", "Pcs.", 75.00),
    (42, "Alu Beams (Light Duty)", "Alu. Beam Light duty -150x 2.0m", "Pcs.", 60.00),

    # Alu. Beam- Heavy duty
    (43, "Alu Beams (Heavy Duty)", "Alu. Beam- heavy duty - 150x 6.0m", "Pcs.", 220.00),
    (44, "Alu Beams (Heavy Duty)", "Alu. Beam Heavy duty -150x5.5m", "Pcs.", 200.00),
    (45, "Alu Beams (Heavy Duty)", "Alu. Beam Heavy duty - 150x5.0m", "Pcs.", 185.00),
    (46, "Alu Beams (Heavy Duty)", "Alu. Beam Heavy duty -150X 4.5m", "Pcs.", 165.00),
    (47, "Alu Beams (Heavy Duty)", "Alu. Beam Heavy duty -150X 4.0m", "Pcs.", 150.00),
    (48, "Alu Beams (Heavy Duty)", "Alu. Beam Heavy duty -150X3.5m", "Pcs.", 130.00),
    (49, "Alu Beams (Heavy Duty)", "Alu. Beam Heavy duty -150X3.0m", "Pcs.", 110.00),
    (50, "Alu Beams (Heavy Duty)", "Alu. Beam Heavy duty -150X2.5m", "Pcs.", 95.00),
    (51, "Alu Beams (Heavy Duty)", "Alu. Beam Heavy duty -150x 2.0m", "Pcs.", 80.00),

    # Cantilever Frames
    (52, "Frames & Cantilever", "Canti liver Frame 1.5m", "Pcs.", 85.00),
    (53, "Frames & Cantilever", "Canti liver Frame 1.3m", "Pcs.", 75.00),

    # Scaffold Tubes
    (54, "Scaffold Tubes", "Scaffold Tube 0.5m", "Pcs.", 5.50),
    (55, "Scaffold Tubes", "Scaffold Tube 1.0m", "Pcs.", 11.00),
    (56, "Scaffold Tubes", "Scaffold Tube 1.5m", "Pcs.", 17.00),
    (57, "Scaffold Tubes", "Scaffold Tube 2.0m", "Pcs.", 24.00),
    (58, "Scaffold Tubes", "Scaffold Tube 2.5m", "Pcs.", 31.00),
    (59, "Scaffold Tubes", "3.0m Scaffold Tube", "Pcs.", 38.00),
    (60, "Scaffold Tubes", "Scaffold Tube 3.5m", "Pcs.", 44.00),
    (61, "Scaffold Tubes", "4.0m Scaffold Tube", "Pcs.", 51.00),
    (62, "Scaffold Tubes", "Scaffold Tube 4.5m", "Pcs.", 57.00),
    (63, "Scaffold Tubes", "Scaffold Tube 5.0m", "Pcs.", 63.00),
    (64, "Scaffold Tubes", "Scaffold Tube 5.5m", "Pcs.", 66.00),
    (65, "Scaffold Tubes", "6.0m Scaffold Tube", "Pcs.", 69.00),

    # Couplers & Clamps
    (66, "Couplers & Clamps", "Swivel Coupler", "Pcs.", 5.10),
    (67, "Couplers & Clamps", "Double Coupler", "Pcs.", 3.25),
    (68, "Couplers & Clamps", "Board Clamp", "Pcs.", 3.50),
    (69, "Couplers & Clamps", "Joint Box(Sleeve Coupler)", "Pcs.", 6.50),
    (70, "Couplers & Clamps", "Joint Pin (Spigot with Nut and Bolt)", "Pcs.", 5.00),
    (71, "Couplers & Clamps", "Girder Clamp", "Pcs.", 8.00),
    (72, "Couplers & Clamps", "Put log Coupler", "Pcs.", 3.50),
    (73, "Couplers & Clamps", "Strairthread Clamp", "Pcs.", 6.00),
    (74, "Couplers & Clamps", "Ladder Clamp", "Pcs.", 6.50),

    # Castor Wheels & Accessories
    (75, "Castor Wheels & Tools", "Heavy Duty Castor Wheel -6\"", "Pcs.", 45.00),
    (76, "Castor Wheels & Tools", "Heavy Duty Castor Wheel -8\"", "Pcs.", 55.00),
    (77, "Castor Wheels & Tools", "Light Duty Castor Wheel -6\"", "Pcs.", 30.00),
    (78, "Castor Wheels & Tools", "Light Duty Castor Wheel -8\"", "Pcs.", 40.00),
    (79, "Couplers & Clamps", "T-Bolts for scaffold Coupler", "Pcs.", 2.00),
    (80, "Castor Wheels & Tools", "Scaffold Spanner", "Pcs.", 25.00),
    (81, "Castor Wheels & Tools", "GIN Wheel-BS 1692;1998", "Pcs.", 45.00),
    (82, "Castor Wheels & Tools", "Spirit Level -10\"", "Pcs.", 20.00),
    (83, "Castor Wheels & Tools", "Scaffolding Lifting Bag", "Pcs.", 35.00),
    (84, "Castor Wheels & Tools", "Scaffolding Tag with Holder", "Pcs.", 15.00),
    (85, "Castor Wheels & Tools", "ScaffoldinTag", "Pcs.", 5.00),
    (86, "Castor Wheels & Tools", "Scaffold Tool Belt", "Pcs.", 40.00),
    (87, "Castor Wheels & Tools", "Spigot Connector", "Pcs.", 6.00),
    (88, "Castor Wheels & Tools", "Scaffold Tube End cap", "Pcs.", 1.50),

    # LVL Boards
    (89, "LVL Boards", "LVL Board 6.0m", "Pcs.", 54.00),
    (90, "LVL Boards", "LVL Board 5.5m", "Pcs.", 49.00),
    (91, "LVL Boards", "LVL Board 5.0m", "Pcs.", 45.00),
    (92, "LVL Boards", "LVL Board 4.5m", "Pcs.", 40.00),
    (93, "LVL Boards", "LVL Board 4.0m", "Pcs.", 36.00),
    (94, "LVL Boards", "LVL Board 3.5m", "Pcs.", 31.00),
    (95, "LVL Boards", "3.0m Scaffold Board(LVL)", "Pcs.", 27.00),
    (96, "LVL Boards", "LVL Board 2.5m", "Pcs.", 25.00),
    (97, "LVL Boards", "2.0m Scaffold Board(LVL)", "Pcs.", 29.00),
    (98, "LVL Boards", "LVL Board 1.5m", "Pcs.", 20.00),
    (99, "LVL Boards", "LVL Board 1.0m", "Pcs.", 15.00),
    (100, "LVL Boards", "LVL Board 0.5m", "Pcs.", 10.00),

    # Steel Boards
    (101, "Steel Boards", "Steel Board 4.0m", "Pcs.", 65.00),
    (102, "Steel Boards", "Steel Board 3.0m", "Pcs.", 50.00),
    (103, "Steel Boards", "Steel Board 2.5m", "Pcs.", 42.00),
    (104, "Steel Boards", "Steel Board 2.0m", "Pcs.", 35.00),

    # Steel Ladders
    (105, "Ladders", "Steel Ladder 6m", "Pcs.", 180.00),
    (106, "Ladders", "Steel Ladder 5m", "Pcs.", 150.00),
    (107, "Ladders", "Steel Ladder  4.0m", "Pcs.", 120.00),
    (108, "Ladders", "Steel Ladder 3m", "Pcs.", 90.00),
    (109, "Ladders", "Steel Ladder 2m", "Pcs.", 60.00),

    # Aluminium Ladders
    (110, "Ladders", "Aluminium Ladder 6.0m", "Pcs.", 240.00),
    (111, "Ladders", "Aluminium Ladder 4.2m", "Pcs.", 170.00),
    (112, "Ladders", "Aluminium Ladder 3.5m", "Pcs.", 140.00),
    (113, "Ladders", "Aluminium Ladder 3.0m", "Pcs.", 120.00),
    (114, "Ladders", "Aluminium Ladder 2.0m", "Pcs.", 80.00),

    # Other Alu Beams & Props
    (115, "Alu Beams & Props", "Alu.Beam 4.0m", "Pcs.", 120.00),
    (116, "Alu Beams & Props", "Alu.Beam 2.5m", "Pcs.", 75.00),
    (117, "Alu Beams & Props", "Alu.Beam 2.2m", "Pcs.", 66.00),
    (118, "Alu Beams & Props", "Alu.Beam 2.4m", "Pcs.", 72.00),
    (119, "Alu Beams & Props", "Alu.Beam 2m", "Pcs.", 60.00),
    (120, "Alu Beams & Props", "Steel Props", "Pcs.", 45.00),
    (121, "Castor Wheels & Tools", "Castor Wheel 8\"", "Pcs.", 55.00),

    # Aluminium Tower Accessories & Frames
    (122, "Towers, Frames & Braces", "Aluminium Tower Accessories", "Pcs.", 50.00),
    (123, "Towers, Frames & Braces", "Caster wheel", "Pcs.", 40.00),
    (124, "Towers, Frames & Braces", "Guard rail", "Pcs.", 35.00),
    (125, "Towers, Frames & Braces", "Stabilazer 6m", "Pcs.", 95.00),
    (126, "Towers, Frames & Braces", "Ladder frame 80cm", "Pcs.", 65.00),
    (127, "Towers, Frames & Braces", "Plain Frame 80cm", "Pcs.", 55.00),
    (128, "Towers, Frames & Braces", "Stabilazer 3m", "Pcs.", 60.00),
    (129, "Towers, Frames & Braces", "Toaboard-wooden", "Pcs.", 25.00),
    (130, "Towers, Frames & Braces", "Ladder frame 140cm", "Pcs.", 85.00),
    (131, "Towers, Frames & Braces", "Plain Frame 140cm", "Pcs.", 75.00),
    (132, "Towers, Frames & Braces", "Platform-Trapdoor", "Pcs.", 110.00),
    (133, "Towers, Frames & Braces", "Platform-Plain", "Pcs.", 85.00),
    (134, "Towers, Frames & Braces", "Horizonatl Brace", "Pcs.", 35.00),
    (135, "Towers, Frames & Braces", "Diagonal brace", "Pcs.", 35.00),
    (136, "Towers, Frames & Braces", "stabilazor 4m", "Pcs.", 75.00)
]

DEFAULT_TERMS = """1. Quotation valid for 7days Only
2. Availability : ex-Stock, Subjected prior sale order
3. Payment along with order
4. Delivery with in 5 days from the confirmation"""

DEFAULT_BANK_DETAILS = """Our : Bank : Saudi National Bank
Account # 04100000530304
IBAN; SA 8910000004100000530304"""

DEFAULT_SIGNATORY = "Mohamed Faizal"
DEFAULT_SIGNATORY_TITLE = "Rawaiya AL Etihad Est."

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db()
    cursor = conn.cursor()

    cursor.execute('''
    CREATE TABLE IF NOT EXISTS PriceList (
        item_id INTEGER PRIMARY KEY AUTOINCREMENT,
        item_number INTEGER,
        category TEXT NOT NULL,
        item_name TEXT NOT NULL,
        unit TEXT DEFAULT 'Pcs.',
        unit_price REAL DEFAULT 0.0,
        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    ''')

    cursor.execute('''
    CREATE TABLE IF NOT EXISTS Projects (
        project_id INTEGER PRIMARY KEY AUTOINCREMENT,
        company_name TEXT NOT NULL,
        site_name TEXT,
        location TEXT,
        attn TEXT,
        tel TEXT,
        email TEXT,
        project_status TEXT DEFAULT 'Lead',
        quotation_ref TEXT,
        quotation_date TEXT,
        subject TEXT DEFAULT 'Scaffolding Materials on sale basis(Used and Refurbed materials)',
        intro_note TEXT DEFAULT 'We thank you for whatsapp inquiry and pleased quote for used equipment as follows:',
        terms_conditions TEXT,
        bank_details TEXT,
        signatory_name TEXT,
        signatory_title TEXT,
        subtotal REAL DEFAULT 0.0,
        vat_amount REAL DEFAULT 0.0,
        grand_total REAL DEFAULT 0.0,
        items_json TEXT DEFAULT '[]',
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    ''')

    cursor.execute('SELECT COUNT(*) FROM PriceList')
    count = cursor.fetchone()[0]
    if count == 0:
        for item_num, cat, name, unit, price in SCAFFOLDING_COMPONENTS:
            cursor.execute('''
            INSERT INTO PriceList (item_number, category, item_name, unit, unit_price)
            VALUES (?, ?, ?, ?, ?)
            ''', (item_num, cat, name, unit, price))
        conn.commit()
        print(f"Successfully initialized database with {len(SCAFFOLDING_COMPONENTS)} items.")
    else:
        print(f"PriceList already has {count} items.")

    conn.close()

def get_next_quotation_ref():
    """
    Computes the next official quotation reference number, continuing the TMR-FZ sequence:
    e.g., TMR-FZ-1745, TMR-FZ-1746, TMR-FZ-1747 -> TMR-FZ-1748-55471-2026
    """
    highest = 1747
    word_dir = r'D:\TAMEER\Quotations\Word'
    if os.path.exists(word_dir):
        for fname in os.listdir(word_dir):
            m = re.search(r'TMR-FZ-(\d+)-55471', fname, re.IGNORECASE)
            if m:
                highest = max(highest, int(m.group(1)))

    if os.path.exists(DB_PATH):
        try:
            conn = sqlite3.connect(DB_PATH)
            cur = conn.cursor()
            cur.execute('SELECT quotation_ref FROM Projects')
            for (ref,) in cur.fetchall():
                if ref:
                    m = re.search(r'TMR-FZ-(\d+)-55471', ref, re.IGNORECASE)
                    if m:
                        highest = max(highest, int(m.group(1)))
            conn.close()
        except:
            pass

    next_seq = highest + 1
    current_year = datetime.now().strftime("%Y")
    return f"TMR-FZ-{next_seq}-55471-{current_year}"

if __name__ == '__main__':
    init_db()

