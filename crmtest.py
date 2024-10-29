import sqlite3
import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinterdnd2 import TkinterDnD, DND_FILES
import os
import platform
import subprocess
import re  # Import regular expressions module

# Database functions
def create_connection():
    conn = sqlite3.connect("crm_leads.db")
    create_leads_table(conn)
    return conn

def create_leads_table(conn):
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS leads (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            Title TEXT,
            Rating TEXT,
            Reviews TEXT,
            Phone TEXT UNIQUE,
            Industry TEXT,
            Address TEXT,
            Website TEXT,
            Google_Maps_Link TEXT,
            Notes TEXT
        )
    """)
    conn.commit()

def load_leads_from_db(conn):
    cursor = conn.cursor()
    cursor.execute("""
        SELECT Title, Rating, Reviews, Phone, Industry, Address, Website, Google_Maps_Link, Notes
        FROM leads
        ORDER BY CASE WHEN Notes IS NOT NULL AND Notes != '' THEN 0 ELSE 1 END, Title
    """)
    rows = cursor.fetchall()
    return pd.DataFrame(rows, columns=[
        'Title', 'Rating', 'Reviews', 'Phone', 'Industry', 'Address',
        'Website', 'Google Maps Link', 'Notes'
    ])

def save_new_leads(conn, new_data):
    cursor = conn.cursor()
    try:
        conn.execute("BEGIN TRANSACTION;")
        for _, row in new_data.iterrows():
            cursor.execute("""
                INSERT OR IGNORE INTO leads (Title, Rating, Reviews, Phone, Industry, Address, Website, Google_Maps_Link, Notes)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                row['Title'], row['Rating'], row['Reviews'], row['Phone'],
                row['Industry'], row['Address'], row['Website'],
                row['Google Maps Link'], row.get('Notes', '')
            ))
        conn.commit()
    except sqlite3.IntegrityError as e:
        print(f"Error saving leads: {e}")
        conn.rollback()

def delete_lead(conn, phone_number):
    cursor = conn.cursor()
    cursor.execute("DELETE FROM leads WHERE Phone = ?", (phone_number,))
    conn.commit()

def update_lead_note(conn, phone_number, note):
    cursor = conn.cursor()
    cursor.execute("UPDATE leads SET Notes = ? WHERE Phone = ?", (note, phone_number))
    conn.commit()

def update_lead(conn, lead_data):
    cursor = conn.cursor()
    cursor.execute("""
        UPDATE leads SET
            Title = ?,
            Rating = ?,
            Reviews = ?,
            Industry = ?,
            Address = ?,
            Website = ?,
            Google_Maps_Link = ?
        WHERE Phone = ?
    """, (
        lead_data['Title'], lead_data['Rating'], lead_data['Reviews'], lead_data['Industry'],
        lead_data['Address'], lead_data['Website'], lead_data['Google Maps Link'], lead_data['Phone']
    ))
    conn.commit()

class CRMClient:
    def __init__(self, root):
        self.root = root  # Use the passed root
        self.root.title("CRM Client")
        self.root.geometry("900x600")
        self.root.overrideredirect(True)  # Remove the default window decorations
        self.root.config(bg='dark grey')
        self.root.rowconfigure(0, weight=1)
        self.root.columnconfigure(0, weight=1)

        # Variables to handle window movement
        self.offset_x = 0
        self.offset_y = 0

        # Create style
        style = ttk.Style()
        style.theme_use('default')

        # Configure styles for ttk widgets
        style.configure('TFrame', background='dark grey')
        style.configure('TLabel', background='dark grey', foreground='white')
        style.configure('TButton', background='grey30', foreground='white')
        style.configure('TEntry', fieldbackground='grey', foreground='white')
        style.configure('TScrollbar', background='dark grey', troughcolor='grey')

        # Configure Treeview style
        style.configure('Custom.Treeview',
                        background='grey20',
                        foreground='white',
                        fieldbackground='grey20',
                        highlightthickness=0,
                        borderwidth=0)
        style.map('Custom.Treeview',
                  background=[('selected', 'grey50')],
                  foreground=[('selected', 'white')])

        style.configure('Custom.Treeview.Heading',
                        background='grey30',
                        foreground='white',
                        relief='flat')

        self.conn = create_connection()
        self.crm_data = load_leads_from_db(self.conn)
        self.root.drop_target_register(DND_FILES)
        self.root.dnd_bind('<<Drop>>', self.handle_drop_event)

        self.setup_ui()
        self.load_leads()

    def setup_ui(self):
        # Create custom title bar
        title_bar = tk.Frame(self.root, bg='grey30', relief='raised', bd=0)
        title_bar.pack(side='top', fill='x')

        # Title bar label
        title_label = tk.Label(title_bar, text="CRM Client", bg='grey30', fg='white', font=("Helvetica", 12))
        title_label.pack(side='left', padx=10)

        # Minimize, Maximize, Close buttons
        btn_close = tk.Button(title_bar, text='✕', command=self.root.destroy, bg='grey30', fg='white', bd=0)
        btn_close.pack(side='right', padx=5)
        btn_maximize = tk.Button(title_bar, text='🗖', command=self.toggle_maximize, bg='grey30', fg='white', bd=0)
        btn_maximize.pack(side='right')
        btn_minimize = tk.Button(title_bar, text='🗕', command=self.minimize_window, bg='grey30', fg='white', bd=0)
        btn_minimize.pack(side='right')

        # Bind events for moving the window
        title_bar.bind('<ButtonPress-1>', self.start_move)
        title_bar.bind('<ButtonRelease-1>', self.stop_move)
        title_bar.bind('<B1-Motion>', self.do_move)

        # Main frame under the title bar
        main_frame = ttk.Frame(self.root, padding=10, style='TFrame')
        main_frame.pack(expand=True, fill='both')
        main_frame.rowconfigure(3, weight=1)
        main_frame.columnconfigure(0, weight=1)

        # Menu bar within the application
        menu_bar_frame = tk.Frame(main_frame, bg='grey30')
        menu_bar_frame.grid(row=0, column=0, sticky='ew')
        menu_bar_frame.columnconfigure(0, weight=1)

        file_menu_button = tk.Menubutton(menu_bar_frame, text='File', bg='grey30', fg='white', activebackground='grey50', activeforeground='white', relief='flat')
        file_menu_button.menu = tk.Menu(file_menu_button, tearoff=0, bg='grey30', fg='white', activebackground='grey50', activeforeground='white')
        file_menu_button['menu'] = file_menu_button.menu
        file_menu_button.menu.add_command(label='Import CSV', command=self.import_csv)
        file_menu_button.pack(side='left', padx=5)

        # Title label within the main frame
        self.title_label = tk.Label(main_frame, text="CRM Client", font=("Helvetica", 18), bg='dark grey', fg='white')
        self.title_label.grid(row=1, column=0, columnspan=3, sticky="w")

        tk.Button(main_frame, text="Add", command=self.add_lead, bg='grey30', fg='white').grid(row=2, column=0, padx=5, pady=5, sticky="w")
        tk.Button(main_frame, text="Delete", command=self.delete_lead, bg='grey30', fg='white').grid(row=2, column=1, padx=5, pady=5, sticky="w")

        table_frame = ttk.Frame(main_frame, style='TFrame')
        table_frame.grid(row=3, column=0, columnspan=3, sticky="nsew")

        self.v_scrollbar = ttk.Scrollbar(table_frame, orient="vertical")
        self.v_scrollbar.grid(row=0, column=1, sticky="ns")

        self.h_scrollbar = ttk.Scrollbar(table_frame, orient="horizontal")
        self.h_scrollbar.grid(row=1, column=0, sticky="ew")

        columns = list(self.crm_data.columns[:-1])  # Exclude 'Notes' column from display
        self.lead_table = ttk.Treeview(
            table_frame, columns=columns, show="headings",
            yscrollcommand=self.v_scrollbar.set, xscrollcommand=self.h_scrollbar.set,
            style='Custom.Treeview'
        )

        for col in columns:
            self.lead_table.heading(col, text=col)
            self.lead_table.column(col, minwidth=100, width=120, stretch=tk.YES)
        self.lead_table.grid(row=0, column=0, sticky="nsew")

        self.v_scrollbar.config(command=self.lead_table.yview)
        self.h_scrollbar.config(command=self.lead_table.xview)

        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)

        self.lead_table.tag_configure("noted", background="lightcoral")
        self.lead_table.bind("<Double-1>", self.edit_cell)
        self.lead_table.bind("<Button-3>", self.show_context_menu)
        self.lead_table.bind("<Button-1>", self.on_click)

        self.context_menu = tk.Menu(self.root, tearoff=0, bg='dark grey', fg='white', activebackground='grey30', activeforeground='white')
        self.context_menu.add_command(label="Add/Edit Note", command=self.add_edit_note)
        self.context_menu.add_command(label="Edit Lead", command=self.edit_lead)

    # Methods for window movement
    def start_move(self, event):
        self.offset_x = event.x
        self.offset_y = event.y

    def stop_move(self, event):
        self.offset_x = None
        self.offset_y = None

    def do_move(self, event):
        x = self.root.winfo_pointerx() - self.offset_x
        y = self.root.winfo_pointery() - self.offset_y
        self.root.geometry(f"+{x}+{y}")

    # Methods for window control buttons
    def minimize_window(self):
        self.root.overrideredirect(False)
        self.root.iconify()
        self.root.update()
        self.root.overrideredirect(True)

    def toggle_maximize(self):
        if self.root.state() == 'zoomed':
            self.root.state('normal')
        else:
            self.root.state('zoomed')

    def handle_drop_event(self, event):
        file_path = event.data.strip('{}')
        if file_path.endswith('.csv'):
            self.import_csv(file_path)

    def import_csv(self, file_path=None):
        if not file_path:
            file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if file_path:
            try:
                new_data = pd.read_csv(file_path)
                required_columns = set(self.crm_data.columns) - {'Notes'}
                if not required_columns.issubset(new_data.columns) or new_data.empty:
                    messagebox.showerror("Import Error", "CSV file format is invalid or empty.")
                    return
                save_new_leads(self.conn, new_data)
                self.crm_data = load_leads_from_db(self.conn)
                self.load_leads()
            except Exception as e:
                messagebox.showerror("Import Error", f"An error occurred while importing the CSV file:\n{e}")

    def load_leads(self):
        for item in self.lead_table.get_children():
            self.lead_table.delete(item)
        for _, row in self.crm_data.iterrows():
            title = f"{row['Title']} *" if row['Notes'] else row['Title']
            tag = "noted" if row['Notes'] else ""
            self.lead_table.insert('', 'end', values=[
                title, row['Rating'], row['Reviews'], row['Phone'], row['Industry'],
                row['Address'], row['Website'], row['Google Maps Link']
            ], tags=(tag,))

    def show_context_menu(self, event):
        item_id = self.lead_table.identify_row(event.y)
        if item_id:
            self.lead_table.selection_set(item_id)
            self.context_menu.post(event.x_root, event.y_root)

    def add_edit_note(self):
        selected_item = self.lead_table.selection()
        if not selected_item:
            messagebox.showerror("No Lead Selected", "Select a lead to add a note.")
            return

        lead_values = self.lead_table.item(selected_item, 'values')
        phone_number = lead_values[3]
        current_note = self.crm_data.loc[self.crm_data['Phone'] == phone_number, 'Notes'].values[0]

        note_window = tk.Toplevel(self.root)
        note_window.title("Add/Edit Note")
        note_window.configure(bg='dark grey')

        tk.Label(note_window, text="Note:", bg='dark grey', fg='white').grid(row=0, column=0, padx=5, pady=5)
        note_text = tk.Text(note_window, width=40, height=10, bg='grey', fg='white', insertbackground='white')
        note_text.insert("1.0", current_note)
        note_text.grid(row=1, column=0, padx=5, pady=5)

        def save_note():
            new_note = note_text.get("1.0", tk.END).strip()
            update_lead_note(self.conn, phone_number, new_note)
            self.crm_data.loc[self.crm_data['Phone'] == phone_number, 'Notes'] = new_note
            self.load_leads()
            note_window.destroy()

        tk.Button(note_window, text="Save", command=save_note, bg='grey30', fg='white').grid(row=2, column=0, padx=5, pady=5)

    def add_lead(self):
        add_window = tk.Toplevel(self.root)
        add_window.title("Add New Lead")
        add_window.configure(bg='dark grey')

        fields = ['Title', 'Rating', 'Reviews', 'Phone', 'Industry', 'Address', 'Website', 'Google Maps Link']
        entries = {}

        for idx, field in enumerate(fields):
            tk.Label(add_window, text=field + ":", bg='dark grey', fg='white').grid(row=idx, column=0, padx=5, pady=5, sticky='e')
            entry = tk.Entry(add_window, bg='grey', fg='white', insertbackground='white')
            entry.grid(row=idx, column=1, padx=5, pady=5)
            entries[field] = entry

        def save_new_lead():
            lead_data = {field: entries[field].get() for field in fields}
            if not lead_data['Phone']:
                messagebox.showerror("Input Error", "Phone number is required.")
                return
            new_data = pd.DataFrame([lead_data])
            try:
                save_new_leads(self.conn, new_data)
                self.crm_data = load_leads_from_db(self.conn)
                self.load_leads()
                add_window.destroy()
            except sqlite3.IntegrityError as e:
                messagebox.showerror("Database Error", f"An error occurred:\n{e}")

        tk.Button(add_window, text="Save", command=save_new_lead, bg='grey30', fg='white').grid(row=len(fields), column=0, columnspan=2, pady=10)

    def delete_lead(self):
        selected_item = self.lead_table.selection()
        if not selected_item:
            messagebox.showerror("No Lead Selected", "Select a lead to delete.")
            return

        lead_values = self.lead_table.item(selected_item, 'values')
        phone_number = lead_values[3]

        # Confirm deletion
        confirm = messagebox.askyesno("Confirm Delete", f"Are you sure you want to delete the lead with phone number {phone_number}?")
        if confirm:
            delete_lead(self.conn, phone_number)
            self.crm_data = load_leads_from_db(self.conn)
            self.load_leads()

    def edit_cell(self, event):
        selected_item = self.lead_table.selection()
        if not selected_item:
            return

        self.edit_lead()

    def edit_lead(self):
        selected_item = self.lead_table.selection()
        if not selected_item:
            messagebox.showerror("No Lead Selected", "Select a lead to edit.")
            return

        lead_values = self.lead_table.item(selected_item, 'values')
        phone_number = lead_values[3]
        lead_row = self.crm_data.loc[self.crm_data['Phone'] == phone_number].iloc[0]

        edit_window = tk.Toplevel(self.root)
        edit_window.title("Edit Lead")
        edit_window.configure(bg='dark grey')

        fields = ['Title', 'Rating', 'Reviews', 'Phone', 'Industry', 'Address', 'Website', 'Google Maps Link']
        entries = {}

        for idx, field in enumerate(fields):
            tk.Label(edit_window, text=field + ":", bg='dark grey', fg='white').grid(row=idx, column=0, padx=5, pady=5, sticky='e')
            entry = tk.Entry(edit_window, bg='grey', fg='white', insertbackground='white')
            # Handle NaN or None values
            value = lead_row[field]
            if pd.isna(value):
                value = ''
            else:
                value = str(value)
            entry.insert(0, value)
            entry.grid(row=idx, column=1, padx=5, pady=5)
            entries[field] = entry

        def save_edited_lead():
            lead_data = {field: entry.get() for field, entry in entries.items()}
            if not lead_data['Phone']:
                messagebox.showerror("Input Error", "Phone number is required.")
                return
            try:
                update_lead(self.conn, lead_data)
                self.crm_data = load_leads_from_db(self.conn)
                self.load_leads()
                edit_window.destroy()
            except sqlite3.IntegrityError as e:
                messagebox.showerror("Database Error", f"An error occurred:\n{e}")

        tk.Button(edit_window, text="Save", command=save_edited_lead, bg='grey30', fg='white').grid(row=len(fields), column=0, columnspan=2, pady=10)

    def on_click(self, event):
        # Identify the item and column
        item_id = self.lead_table.identify_row(event.y)
        col = self.lead_table.identify_column(event.x)
        col_index = int(col[1:]) - 1  # Convert from '#1' to 0

        if item_id and col_index == 3:  # If 'Phone' column is clicked
            lead_values = self.lead_table.item(item_id, 'values')
            phone_number = lead_values[3]
            self.call_with_skype_uri(phone_number)

    def call_with_skype_uri(self, phone_number):
        if phone_number:
            try:
                # Remove non-numeric characters
                cleaned_number = re.sub(r'\D', '', phone_number)
                # Optionally remove leading '1' if present and number is 11 digits
                if cleaned_number.startswith('1') and len(cleaned_number) == 11:
                    cleaned_number = cleaned_number[1:]
                skype_uri = f"callto://{cleaned_number}"
                system_platform = platform.system()
                if system_platform == 'Windows':
                    os.startfile(skype_uri)
                elif system_platform == 'Darwin':  # macOS
                    subprocess.Popen(['open', skype_uri])
                else:  # Assume Linux
                    subprocess.Popen(['xdg-open', skype_uri])
            except Exception as e:
                messagebox.showerror("Error", f"Failed to call using Skype.\n{e}")

# Run the application
if __name__ == "__main__":
    root = TkinterDnD.Tk()
    app = CRMClient(root)
    root.mainloop()
