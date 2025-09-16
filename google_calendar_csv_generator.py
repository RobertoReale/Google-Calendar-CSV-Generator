import csv
import json
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, scrolledtext
from datetime import datetime, timedelta
from typing import List, Dict, Optional
import os
from dataclasses import dataclass
from enum import Enum
from collections import defaultdict

class DateFormat(Enum):
    ITALIAN = "DD/MM/YYYY"
    ISO = "YYYY-MM-DD"

WEEKDAYS_IT = ["Lunedì", "Martedì", "Mercoledì", "Giovedì", "Venerdì", "Sabato", "Domenica"]

@dataclass
class Session:
    weekday: int  # 0=Lunedì, 6=Domenica
    start_time: str  # HH:MM
    end_time: str  # HH:MM
    location: str
    
    def __str__(self):
        return f"{WEEKDAYS_IT[self.weekday]} {self.start_time}-{self.end_time} @ {self.location}"

@dataclass
class CourseEvent:
    subject: str
    description: str
    start_date: str
    end_date: str
    sessions: List[Session]
    is_private: bool = True
    all_day: bool = False
    
    def get_total_occurrences(self) -> int:
        if not self.sessions:
            return 0
        try:
            start = self.parse_date(self.start_date)
            end = self.parse_date(self.end_date)
            weeks = ((end - start).days // 7) + 1
            return len(self.sessions) * weeks
        except:
            return 0

    @staticmethod
    def parse_date(date_str: str) -> datetime:
        for fmt in ["%d/%m/%Y", "%Y-%m-%d"]:
            try:
                return datetime.strptime(date_str, fmt)
            except:
                continue
        raise ValueError(f"Formato data non valido: {date_str}")

def parse_date_flexible(date_str: str) -> Optional[datetime]:
    date_str = date_str.strip()
    for fmt in ["%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%Y/%m/%d", "%m/%d/%Y"]:
        try:
            return datetime.strptime(date_str, fmt)
        except:
            continue
    return None

def parse_time(time_str: str) -> Optional[datetime]:
    try:
        return datetime.strptime(time_str.strip(), "%H:%M")
    except:
        return None

def parse_time_12h(time_str: str) -> str:
    try:
        if not time_str or time_str.strip() == "":
            return "09:00"
        time_str = time_str.strip()
        time_obj = datetime.strptime(time_str, "%I:%M %p")
        return time_obj.strftime("%H:%M")
    except:
        try:
            if time_str.startswith("0"):
                time_str = time_str[1:]
            time_obj = datetime.strptime(time_str, "%I:%M %p")
            return time_obj.strftime("%H:%M")
        except:
            return "09:00"

def format_date_for_google(date: datetime) -> str:
    return date.strftime("%m/%d/%Y")

def format_time_12h(time_str: str) -> str:
    try:
        time_obj = datetime.strptime(time_str, "%H:%M")
        return time_obj.strftime("%I:%M %p").lstrip("0")
    except:
        return "09:00 AM"

def get_dates_for_weekday(start_date: datetime, end_date: datetime, weekday: int) -> List[datetime]:
    current = start_date
    days_ahead = (weekday - current.weekday()) % 7
    current = current + timedelta(days=days_ahead)
    dates = []
    while current <= end_date:
        dates.append(current)
        current += timedelta(weeks=1)
    return dates

def load_csv_events(filename: str) -> List[CourseEvent]:
    """Carica eventi da un file CSV e li converte in CourseEvent"""
    events = []
    
    try:
        with open(filename, 'r', encoding='utf-8') as csvfile:
            try:
                content = csvfile.read()
            except UnicodeDecodeError:
                csvfile.close()
                with open(filename, 'r', encoding='latin1') as csvfile2:
                    content = csvfile2.read()
            
            lines = content.splitlines()
            reader = csv.DictReader(lines)
            
            # Group events by subject and description
            event_groups = defaultdict(list)
            
            for row in reader:
                subject = row.get('Subject', '').strip()
                description = row.get('Description', '').strip()
                
                if not subject:
                    continue
                
                key = f"{subject}|{description}"
                event_groups[key].append(row)
            
            # Convert groups back to CourseEvent objects
            for key, rows in event_groups.items():
                subject, description = key.split('|', 1)
                events.append(csv_rows_to_course_event(subject, description, rows))
        
        return events
        
    except Exception as e:
        raise Exception(f"Errore nel caricamento del CSV: {str(e)}")

def csv_rows_to_course_event(subject: str, description: str, rows: List[Dict]) -> CourseEvent:
    """Converte le righe CSV in un oggetto CourseEvent - VERSIONE CORRETTA E DEBUGGATA"""
    
    first_row = rows[0]
    all_day = first_row.get('All Day Event', 'False').strip().lower() == 'true'
    is_private = first_row.get('Private', 'True').strip().lower() == 'true'
    
    # STEP 1: Parse all dates correctly (CSV uses MM/DD/YYYY format)
    all_dates = []
    for row in rows:
        start_date_str = row.get('Start Date', '').strip()
        if start_date_str:
            # Force MM/DD/YYYY parsing for Google Calendar CSVs
            try:
                date_obj = datetime.strptime(start_date_str, "%m/%d/%Y")
                all_dates.append(date_obj)
            except:
                # Fallback to flexible parsing
                date_obj = parse_date_flexible(start_date_str)
                if date_obj:
                    all_dates.append(date_obj)
    
    if not all_dates:
        start_date = datetime.now().strftime("%d/%m/%Y")
        end_date = (datetime.now() + timedelta(days=90)).strftime("%d/%m/%Y")
    else:
        # Convert to DD/MM/YYYY format for display
        start_date = min(all_dates).strftime("%d/%m/%Y")
        end_date = max(all_dates).strftime("%d/%m/%Y")
    
    sessions = []
    
    if not all_day:
        # STEP 2: Extract unique session patterns
        unique_sessions = set()
        
        for row in rows:
            start_date_str = row.get('Start Date', '').strip()
            start_time_str = row.get('Start Time', '').strip()
            end_time_str = row.get('End Time', '').strip()
            location = row.get('Location', '').strip()
            
            if start_date_str and start_time_str and end_time_str:
                # Parse date with MM/DD/YYYY format
                try:
                    date_obj = datetime.strptime(start_date_str, "%m/%d/%Y")
                except:
                    date_obj = parse_date_flexible(start_date_str)
                
                if date_obj:
                    weekday = date_obj.weekday()  # 0=Monday, 6=Sunday
                    start_time_24h = parse_time_12h(start_time_str)
                    end_time_24h = parse_time_12h(end_time_str)
                    
                    # Normalize location (remove extra spaces)
                    location = ' '.join(location.split())
                    
                    # Create unique session key
                    session_key = (weekday, start_time_24h, end_time_24h, location)
                    unique_sessions.add(session_key)
        
        # STEP 3: Convert unique sessions to Session objects
        for weekday, start_time, end_time, location in sorted(unique_sessions):
            session = Session(
                weekday=weekday,
                start_time=start_time,
                end_time=end_time,
                location=location
            )
            sessions.append(session)
    
    return CourseEvent(
        subject=subject,
        description=description,
        start_date=start_date,
        end_date=end_date,
        sessions=sessions,
        is_private=is_private,
        all_day=all_day
    )

class SessionDialog(tk.Toplevel):
    """Dialog per aggiungere/modificare una sessione"""
    def __init__(self, parent, session: Optional[Session] = None):
        super().__init__(parent)
        self.title("Sessione")
        self.geometry("400x250")
        self.result = None
        self.transient(parent)
        self.grab_set()
        
        frame = ttk.Frame(self, padding="20")
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Form fields
        ttk.Label(frame, text="Giorno:").grid(row=0, column=0, sticky="w", pady=5)
        self.weekday_var = tk.StringVar()
        self.weekday_combo = ttk.Combobox(frame, textvariable=self.weekday_var, 
                                          values=WEEKDAYS_IT, state="readonly", width=20)
        self.weekday_combo.grid(row=0, column=1, pady=5, padx=5)
        
        ttk.Label(frame, text="Ora inizio (HH:MM):").grid(row=1, column=0, sticky="w", pady=5)
        self.start_time_var = tk.StringVar()
        ttk.Entry(frame, textvariable=self.start_time_var, width=22).grid(row=1, column=1, pady=5, padx=5)
        
        ttk.Label(frame, text="Ora fine (HH:MM):").grid(row=2, column=0, sticky="w", pady=5)
        self.end_time_var = tk.StringVar()
        ttk.Entry(frame, textvariable=self.end_time_var, width=22).grid(row=2, column=1, pady=5, padx=5)
        
        ttk.Label(frame, text="Aula/Luogo:").grid(row=3, column=0, sticky="w", pady=5)
        self.location_var = tk.StringVar()
        ttk.Entry(frame, textvariable=self.location_var, width=22).grid(row=3, column=1, pady=5, padx=5)
        
        # Buttons
        button_frame = ttk.Frame(frame)
        button_frame.grid(row=4, column=0, columnspan=2, pady=20)
        
        ttk.Button(button_frame, text="OK", command=self.ok_clicked).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Annulla", command=self.cancel_clicked).pack(side=tk.LEFT, padx=5)
        
        # Load existing session
        if session:
            self.weekday_combo.set(WEEKDAYS_IT[session.weekday])
            self.start_time_var.set(session.start_time)
            self.end_time_var.set(session.end_time)
            self.location_var.set(session.location)
        else:
            self.weekday_combo.current(0)
            self.start_time_var.set("09:00")
            self.end_time_var.set("11:00")
    
    def ok_clicked(self):
        weekday = self.weekday_combo.current()
        if weekday < 0:
            messagebox.showerror("Errore", "Seleziona un giorno")
            return
        
        start_time = self.start_time_var.get().strip()
        end_time = self.end_time_var.get().strip()
        
        if not parse_time(start_time) or not parse_time(end_time):
            messagebox.showerror("Errore", "Formato orario non valido. Usa HH:MM")
            return
        
        if parse_time(end_time) <= parse_time(start_time):
            messagebox.showerror("Errore", "L'ora di fine deve essere dopo l'ora di inizio")
            return
        
        self.result = Session(
            weekday=weekday,
            start_time=start_time,
            end_time=end_time,
            location=self.location_var.get().strip()
        )
        self.destroy()
    
    def cancel_clicked(self):
        self.destroy()

class GoogleCalendarGenerator:
    def __init__(self, root):
        self.root = root
        self.root.title("📅 Google Calendar CSV Generator & Editor")
        self.root.geometry("1200x700")
        self.events: List[CourseEvent] = []
        self.current_sessions: List[Session] = []
        self.date_format = DateFormat.ITALIAN
        self.loaded_csv_filename = None
        
        self.create_menu()
        self.create_widgets()
        self.update_status("Pronto")
    
    def create_menu(self):
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="📂 Carica configurazione", command=self.load_config)
        file_menu.add_command(label="💾 Salva configurazione", command=self.save_config)
        file_menu.add_separator()
        file_menu.add_command(label="📋 Carica CSV esistente", command=self.load_csv)
        file_menu.add_command(label="📊 Genera CSV", command=self.generate_csv)
        file_menu.add_separator()
        file_menu.add_command(label="❌ Esci", command=self.root.quit)
        
        edit_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Modifica", menu=edit_menu)
        edit_menu.add_command(label="➕ Aggiungi corso", command=self.add_course)
        edit_menu.add_command(label="🔄 Reset tutto", command=self.reset_all)
    
    def create_widgets(self):
        main_container = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        main_container.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        left_frame = ttk.Frame(main_container)
        right_frame = ttk.Frame(main_container)
        main_container.add(left_frame, weight=1)
        main_container.add(right_frame, weight=1)
        
        # LEFT PANEL - Course input
        ttk.Label(left_frame, text="📚 Nuovo Corso", font=('Arial', 12, 'bold')).pack(anchor=tk.W, padx=10, pady=5)
        
        details_frame = ttk.LabelFrame(left_frame, text="Dettagli", padding="10")
        details_frame.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Label(details_frame, text="Nome corso:").grid(row=0, column=0, sticky="w", pady=3)
        self.subject_var = tk.StringVar()
        ttk.Entry(details_frame, textvariable=self.subject_var, width=40).grid(row=0, column=1, pady=3)
        
        ttk.Label(details_frame, text="Docente:").grid(row=1, column=0, sticky="w", pady=3)
        self.description_var = tk.StringVar()
        ttk.Entry(details_frame, textvariable=self.description_var, width=40).grid(row=1, column=1, pady=3)
        
        ttk.Label(details_frame, text="Data inizio:").grid(row=2, column=0, sticky="w", pady=3)
        self.start_date_var = tk.StringVar(value="15/09/2025")
        ttk.Entry(details_frame, textvariable=self.start_date_var, width=15).grid(row=2, column=1, pady=3)
        
        ttk.Label(details_frame, text="Data fine:").grid(row=3, column=0, sticky="w", pady=3)
        self.end_date_var = tk.StringVar(value="22/12/2025")
        ttk.Entry(details_frame, textvariable=self.end_date_var, width=15).grid(row=3, column=1, pady=3)
        
        options_frame = ttk.Frame(details_frame)
        options_frame.grid(row=4, column=0, columnspan=2, pady=10)
        
        self.private_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="Privato", variable=self.private_var).pack(side=tk.LEFT, padx=5)
        
        self.all_day_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(options_frame, text="Tutto il giorno", variable=self.all_day_var).pack(side=tk.LEFT, padx=5)
        
        # Sessions
        sessions_frame = ttk.LabelFrame(left_frame, text="Sessioni", padding="10")
        sessions_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        sessions_toolbar = ttk.Frame(sessions_frame)
        sessions_toolbar.pack(fill=tk.X, pady=(0, 5))
        
        ttk.Button(sessions_toolbar, text="➕ Aggiungi", command=self.add_session).pack(side=tk.LEFT, padx=2)
        ttk.Button(sessions_toolbar, text="✏️ Modifica", command=self.edit_session).pack(side=tk.LEFT, padx=2)
        ttk.Button(sessions_toolbar, text="🗑️ Rimuovi", command=self.remove_session).pack(side=tk.LEFT, padx=2)
        
        self.sessions_listbox = tk.Listbox(sessions_frame, height=6)
        self.sessions_listbox.pack(fill=tk.BOTH, expand=True)
        
        add_frame = ttk.Frame(left_frame)
        add_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Button(add_frame, text="➕ AGGIUNGI CORSO", 
                  command=self.add_course).pack(side=tk.LEFT, padx=5)
        ttk.Button(add_frame, text="🔄 Pulisci", command=self.clear_form).pack(side=tk.LEFT, padx=5)
        
        # RIGHT PANEL - Events list
        header_frame = ttk.Frame(right_frame)
        header_frame.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Label(header_frame, text="📚 Eventi", font=('Arial', 12, 'bold')).pack(side=tk.LEFT)
        
        self.csv_indicator = ttk.Label(header_frame, text="", foreground="blue")
        self.csv_indicator.pack(side=tk.RIGHT)
        
        # TreeView
        events_frame = ttk.Frame(right_frame)
        events_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        columns = ('subject', 'docente', 'periodo', 'sessioni', 'eventi')
        self.events_tree = ttk.Treeview(events_frame, columns=columns, show='headings', height=12)
        
        self.events_tree.heading('subject', text='Corso')
        self.events_tree.heading('docente', text='Docente')
        self.events_tree.heading('periodo', text='Periodo')
        self.events_tree.heading('sessioni', text='Sessioni')
        self.events_tree.heading('eventi', text='Eventi')
        
        self.events_tree.column('subject', width=150)
        self.events_tree.column('docente', width=120)
        self.events_tree.column('periodo', width=120)
        self.events_tree.column('sessioni', width=60, anchor=tk.CENTER)
        self.events_tree.column('eventi', width=60, anchor=tk.CENTER)
        
        self.events_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        tree_scrollbar = ttk.Scrollbar(events_frame, orient=tk.VERTICAL)
        tree_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.events_tree.config(yscrollcommand=tree_scrollbar.set)
        tree_scrollbar.config(command=self.events_tree.yview)
        
        # Buttons
        buttons_frame = ttk.Frame(right_frame)
        buttons_frame.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Button(buttons_frame, text="📋 Carica CSV", command=self.load_csv).pack(side=tk.LEFT, padx=2)
        ttk.Button(buttons_frame, text="🔍 Dettagli", command=self.show_event_details).pack(side=tk.LEFT, padx=2)
        ttk.Button(buttons_frame, text="✏️ Modifica", command=self.edit_event).pack(side=tk.LEFT, padx=2)
        ttk.Button(buttons_frame, text="🗑️ Rimuovi", command=self.remove_event).pack(side=tk.LEFT, padx=2)
        ttk.Button(buttons_frame, text="👁️ Anteprima", command=self.preview_csv).pack(side=tk.LEFT, padx=2)
        
        # Statistics
        stats_frame = ttk.LabelFrame(right_frame, text="📊 Statistiche", padding="10")
        stats_frame.pack(fill=tk.X, padx=10, pady=5)
        
        self.stats_label = ttk.Label(stats_frame, text="Nessun evento")
        self.stats_label.pack(anchor=tk.W)
        
        # Generate button
        self.generate_button = ttk.Button(right_frame, text="📊 GENERA CSV", 
                                         command=self.generate_csv)
        self.generate_button.pack(fill=tk.X, padx=10, pady=10)
        
        # Status bar
        self.status_var = tk.StringVar(value="Pronto")
        self.status_bar = ttk.Label(self.root, textvariable=self.status_var, 
                                   relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)
    
    # Session management
    def add_session(self):
        dialog = SessionDialog(self.root)
        self.root.wait_window(dialog)
        if dialog.result:
            self.current_sessions.append(dialog.result)
            self.sessions_listbox.insert(tk.END, str(dialog.result))
            self.update_status(f"Sessione aggiunta")
    
    def edit_session(self):
        selection = self.sessions_listbox.curselection()
        if not selection:
            messagebox.showwarning("Attenzione", "Seleziona una sessione")
            return
        
        idx = selection[0]
        session = self.current_sessions[idx]
        
        dialog = SessionDialog(self.root, session)
        self.root.wait_window(dialog)
        
        if dialog.result:
            self.current_sessions[idx] = dialog.result
            self.sessions_listbox.delete(idx)
            self.sessions_listbox.insert(idx, str(dialog.result))
            self.update_status("Sessione modificata")
    
    def remove_session(self):
        selection = self.sessions_listbox.curselection()
        if not selection:
            return
        for idx in reversed(selection):
            del self.current_sessions[idx]
            self.sessions_listbox.delete(idx)
        self.update_status("Sessione rimossa")
    
    # Event management
    def add_course(self):
        subject = self.subject_var.get().strip()
        if not subject:
            messagebox.showerror("Errore", "Inserisci il nome del corso")
            return
        
        start_date = self.start_date_var.get().strip()
        end_date = self.end_date_var.get().strip()
        
        if not parse_date_flexible(start_date) or not parse_date_flexible(end_date):
            messagebox.showerror("Errore", "Formato data non valido")
            return
        
        if parse_date_flexible(end_date) < parse_date_flexible(start_date):
            messagebox.showerror("Errore", "Data fine deve essere dopo data inizio")
            return
        
        event = CourseEvent(
            subject=subject,
            description=self.description_var.get().strip(),
            start_date=start_date,
            end_date=end_date,
            sessions=self.current_sessions.copy(),
            is_private=self.private_var.get(),
            all_day=self.all_day_var.get()
        )
        
        self.events.append(event)
        self.refresh_events_tree()
        self.update_statistics()
        self.clear_form()
        self.update_status(f"Corso '{subject}' aggiunto")
    
    def edit_event(self):
        selection = self.events_tree.selection()
        if not selection:
            messagebox.showwarning("Attenzione", "Seleziona un evento")
            return

        idx = int(selection[0])
        event = self.events[idx]

        # Load into form
        self.subject_var.set(event.subject)
        self.description_var.set(event.description)
        self.start_date_var.set(event.start_date)
        self.end_date_var.set(event.end_date)
        self.private_var.set(event.is_private)
        self.all_day_var.set(event.all_day)

        self.sessions_listbox.delete(0, tk.END)
        self.current_sessions.clear()
        for s in event.sessions:
            self.current_sessions.append(s)
            self.sessions_listbox.insert(tk.END, str(s))

        # Remove event (will be recreated when saved)
        del self.events[idx]
        self.refresh_events_tree()
        self.update_statistics()
        self.update_status(f"Modifica '{event.subject}' - aggiorna e premi 'Aggiungi Corso'")

    def remove_event(self):
        selection = self.events_tree.selection()
        if not selection:
            messagebox.showwarning("Attenzione", "Seleziona un evento")
            return
        
        idx = int(selection[0])
        event = self.events[idx]
        
        if messagebox.askyesno("Conferma", f"Rimuovere '{event.subject}'?"):
            del self.events[idx]
            self.refresh_events_tree()
            self.update_statistics()
            self.update_status(f"Evento rimosso")
    
    def show_event_details(self):
        selection = self.events_tree.selection()
        if not selection:
            messagebox.showwarning("Attenzione", "Seleziona un evento")
            return
        
        idx = int(selection[0])
        event = self.events[idx]
        
        details = f"📚 {event.subject}\n"
        details += f"👤 {event.description}\n" if event.description else ""
        details += f"📅 {event.start_date} → {event.end_date}\n"
        details += f"🔒 Privato: {'Sì' if event.is_private else 'No'}\n"
        details += f"🌅 Tutto il giorno: {'Sì' if event.all_day else 'No'}\n\n"
        details += "📍 Sessioni:\n"
        
        for session in event.sessions:
            details += f"  • {session}\n"
        
        details += f"\n📊 Totale eventi: {event.get_total_occurrences()}"
        
        messagebox.showinfo("Dettagli", details)
    
    def refresh_events_tree(self):
        for item in self.events_tree.get_children():
            self.events_tree.delete(item)
        
        for idx, event in enumerate(self.events):
            periodo = f"{event.start_date[:10]} → {event.end_date[:10]}"
            self.events_tree.insert('', tk.END, 
                                   iid=str(idx),
                                   values=(event.subject,
                                          event.description,
                                          periodo,
                                          len(event.sessions),
                                          event.get_total_occurrences()))
    
    def update_statistics(self):
        if not self.events:
            self.stats_label.config(text="Nessun evento")
            return
        
        total_events = len(self.events)
        total_sessions = sum(len(e.sessions) for e in self.events)
        total_occurrences = sum(e.get_total_occurrences() for e in self.events)
        
        stats_text = f"Corsi: {total_events} | Sessioni: {total_sessions} | Eventi CSV: {total_occurrences}"
        self.stats_label.config(text=stats_text)
    
    # CSV operations
    def load_csv(self):
        filename = filedialog.askopenfilename(
            title="Carica CSV",
            filetypes=[('CSV files', '*.csv'), ('All files', '*.*')]
        )
        
        if not filename:
            return
        
        try:
            if self.events:
                choice = messagebox.askyesnocancel(
                    "Eventi esistenti",
                    "Sostituire (Sì) o unire (No) agli eventi correnti?"
                )
                if choice is None:
                    return
                elif choice:
                    self.events.clear()
            
            csv_events = load_csv_events(filename)
            
            if not csv_events:
                messagebox.showwarning("Attenzione", "Nessun evento nel CSV")
                return
            
            self.events.extend(csv_events)
            self.loaded_csv_filename = filename
            
            self.refresh_events_tree()
            self.update_statistics()
            self.update_csv_indicator()
            
            messagebox.showinfo("Successo", 
                               f"CSV caricato: {len(csv_events)} eventi")
            self.update_status(f"CSV caricato: {len(csv_events)} eventi")
            
        except Exception as e:
            messagebox.showerror("Errore", f"Errore CSV: {str(e)}")
    
    def update_csv_indicator(self):
        if self.loaded_csv_filename:
            filename = os.path.basename(self.loaded_csv_filename)
            self.csv_indicator.config(text=f"📋 {filename}")
        else:
            self.csv_indicator.config(text="")
    
    def preview_csv(self):
        if not self.events:
            messagebox.showwarning("Attenzione", "Nessun evento")
            return
        
        preview_window = tk.Toplevel(self.root)
        preview_window.title("Anteprima CSV")
        preview_window.geometry("800x500")
        
        text_widget = scrolledtext.ScrolledText(preview_window, wrap=tk.NONE)
        text_widget.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        rows = self.generate_csv_rows()
        
        headers = ['Subject', 'Start Date', 'Start Time', 'End Date', 'End Time', 
                  'All Day Event', 'Description', 'Location', 'Private']
        text_widget.insert(tk.END, ','.join(headers) + '\n')
        text_widget.insert(tk.END, '-' * 100 + '\n')
        
        for i, row in enumerate(rows[:20]):
            line = ','.join([str(row.get(h, '')) for h in headers])
            text_widget.insert(tk.END, f"{i+1}. {line}\n")
        
        if len(rows) > 20:
            text_widget.insert(tk.END, f"\n... e altri {len(rows) - 20} eventi\n")
        
        text_widget.insert(tk.END, f"\nTOTALE: {len(rows)} eventi")
        text_widget.config(state=tk.DISABLED)
        
        ttk.Button(preview_window, text="Chiudi", 
                  command=preview_window.destroy).pack(pady=10)
    
    def generate_csv_rows(self) -> List[Dict]:
        rows = []
        
        for event in self.events:
            if event.all_day:
                start = parse_date_flexible(event.start_date)
                end = parse_date_flexible(event.end_date)
                current = start
                
                while current <= end:
                    row = {
                        'Subject': event.subject,
                        'Start Date': format_date_for_google(current),
                        'Start Time': '',
                        'End Date': format_date_for_google(current),
                        'End Time': '',
                        'All Day Event': 'True',
                        'Description': event.description,
                        'Location': '',
                        'Private': 'True' if event.is_private else 'False'
                    }
                    rows.append(row)
                    current += timedelta(days=1)
            else:
                start = parse_date_flexible(event.start_date)
                end = parse_date_flexible(event.end_date)
                
                for session in event.sessions:
                    dates = get_dates_for_weekday(start, end, session.weekday)
                    
                    for date in dates:
                        row = {
                            'Subject': event.subject,
                            'Start Date': format_date_for_google(date),
                            'Start Time': format_time_12h(session.start_time),
                            'End Date': format_date_for_google(date),
                            'End Time': format_time_12h(session.end_time),
                            'All Day Event': 'False',
                            'Description': event.description,
                            'Location': session.location,
                            'Private': 'True' if event.is_private else 'False'
                        }
                        rows.append(row)
        
        rows.sort(key=lambda r: (r['Start Date'], r['Start Time'] or '00:00'))
        return rows
    
    def generate_csv(self):
        if not self.events:
            messagebox.showerror("Errore", "Nessun evento")
            return
        
        if self.loaded_csv_filename:
            suggested_name = os.path.splitext(self.loaded_csv_filename)[0] + "_modificato.csv"
        else:
            suggested_name = f'calendar_{datetime.now().strftime("%Y%m%d_%H%M%S")}.csv'
        
        filename = filedialog.asksaveasfilename(
            defaultextension='.csv',
            filetypes=[('CSV files', '*.csv')],
            initialfile=os.path.basename(suggested_name)
        )
        
        if not filename:
            return
        
        try:
            rows = self.generate_csv_rows()
            
            with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
                fieldnames = ['Subject', 'Start Date', 'Start Time', 'End Date', 'End Time',
                            'All Day Event', 'Description', 'Location', 'Private']
                writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                writer.writeheader()
                writer.writerows(rows)
            
            messagebox.showinfo("Successo", 
                               f"CSV salvato: {os.path.basename(filename)}\n{len(rows)} eventi")
            self.update_status(f"CSV salvato: {len(rows)} eventi")
            
        except Exception as e:
            messagebox.showerror("Errore", f"Errore salvataggio: {str(e)}")
    
    # Configuration
    def save_config(self):
        if not self.events:
            messagebox.showwarning("Attenzione", "Nessun evento")
            return
        
        filename = filedialog.asksaveasfilename(
            defaultextension='.json',
            filetypes=[('JSON files', '*.json')],
            initialfile='config.json'
        )
        
        if not filename:
            return
        
        try:
            config = {
                'events': []
            }
            
            for event in self.events:
                event_dict = {
                    'subject': event.subject,
                    'description': event.description,
                    'start_date': event.start_date,
                    'end_date': event.end_date,
                    'is_private': event.is_private,
                    'all_day': event.all_day,
                    'sessions': [
                        {
                            'weekday': s.weekday,
                            'start_time': s.start_time,
                            'end_time': s.end_time,
                            'location': s.location
                        }
                        for s in event.sessions
                    ]
                }
                config['events'].append(event_dict)
            
            with open(filename, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=2)
            
            messagebox.showinfo("Successo", "Configurazione salvata")
            self.update_status("Configurazione salvata")
            
        except Exception as e:
            messagebox.showerror("Errore", f"Errore: {str(e)}")
    
    def load_config(self):
        filename = filedialog.askopenfilename(
            filetypes=[('JSON files', '*.json')]
        )
        
        if not filename:
            return
        
        try:
            with open(filename, 'r', encoding='utf-8') as f:
                config = json.load(f)
            
            self.events.clear()
            
            for event_dict in config['events']:
                sessions = []
                for s_dict in event_dict.get('sessions', []):
                    sessions.append(Session(
                        weekday=s_dict['weekday'],
                        start_time=s_dict['start_time'],
                        end_time=s_dict['end_time'],
                        location=s_dict['location']
                    ))
                
                event = CourseEvent(
                    subject=event_dict['subject'],
                    description=event_dict['description'],
                    start_date=event_dict['start_date'],
                    end_date=event_dict['end_date'],
                    sessions=sessions,
                    is_private=event_dict.get('is_private', True),
                    all_day=event_dict.get('all_day', False)
                )
                self.events.append(event)
            
            self.refresh_events_tree()
            self.update_statistics()
            messagebox.showinfo("Successo", "Configurazione caricata")
            self.update_status("Configurazione caricata")
            
        except Exception as e:
            messagebox.showerror("Errore", f"Errore: {str(e)}")
    
    def clear_form(self):
        self.subject_var.set("")
        self.description_var.set("")
        self.sessions_listbox.delete(0, tk.END)
        self.current_sessions.clear()
    
    def reset_all(self):
        if self.events and not messagebox.askyesno("Conferma", "Reset tutto?"):
            return
        
        self.events.clear()
        self.current_sessions.clear()
        self.loaded_csv_filename = None
        self.clear_form()
        self.start_date_var.set("15/09/2025")
        self.end_date_var.set("22/12/2025")
        self.refresh_events_tree()
        self.update_statistics()
        self.update_csv_indicator()
        self.update_status("Reset completato")
    
    def update_status(self, message: str):
        self.status_var.set(message)
        self.root.update_idletasks()

def main():
    root = tk.Tk()
    app = GoogleCalendarGenerator(root)
    
    # Center window
    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f'{width}x{height}+{x}+{y}')
    
    # Shortcuts
    root.bind('<Control-g>', lambda e: app.generate_csv())
    root.bind('<Control-s>', lambda e: app.save_config())
    root.bind('<Control-o>', lambda e: app.load_config())
    root.bind('<Control-l>', lambda e: app.load_csv())
    
    root.mainloop()

if __name__ == '__main__':
    main()