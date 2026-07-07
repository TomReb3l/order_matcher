from __future__ import annotations

import traceback
from datetime import datetime
from pathlib import Path
from typing import Optional

import customtkinter as ctk
import pandas as pd
from tkinter import filedialog, messagebox, ttk

from core import MatcherEngine, PromotionPdfParser, ServiceExcelLoader, WordExporter, TABLE_COLUMNS

APP_TITLE = "Σύγκριση Διαταγής με Δύναμη Υπηρεσίας"
APP_GEOMETRY = "1380x860"
THEME_BLUE = ("#1f6aa5", "#1f6aa5")
APP_DIR = Path(__file__).resolve().parent
ICON_ICO = APP_DIR / "app_icon.ico"
ICON_PNG = APP_DIR / "assets_icon.png"

class ResultsTable(ttk.Treeview):
    def __init__(self, master, columns):
        super().__init__(master, columns=columns, show="headings", height=20)
        self._configure_columns(columns)

    def _configure_columns(self, columns):
        widths = {
            "Α/Α PDF": 80,
            "Μητρώο": 100,
            "Επώνυμο": 180,
            "Όνομα": 160,
            "Πατρώνυμο": 150,
            "Βαθμός": 90,
            "Οργανική": 330,
            "Τρόπος Ταύτισης": 160,
        }
        for col in columns:
            self.heading(col, text=col)
            anchor = "center" if col in ("Α/Α PDF", "Μητρώο", "Βαθμός") else "w"
            self.column(col, width=widths.get(col, 120), anchor=anchor, stretch=True)

    def set_rows(self, df: pd.DataFrame):
        self.delete(*self.get_children())
        if df.empty:
            return
        for _, row in df.iterrows():
            self.insert(
                "",
                "end",
                values=[row.get(col, "") for col in TABLE_COLUMNS],
            )


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry(APP_GEOMETRY)
        self.minsize(1180, 720)
        self._set_window_icon()

        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        self.pdf_path: Optional[str] = None
        self.excel_path: Optional[str] = None
        self.results_bundle: Optional[dict] = None

        self.pdf_parser = PromotionPdfParser()
        self.excel_loader = ServiceExcelLoader()
        self.matcher = MatcherEngine()
        self.word_exporter = WordExporter()

        self._build_ui()

    def _set_window_icon(self):
        try:
            if ICON_ICO.exists():
                self.iconbitmap(str(ICON_ICO))
                return
        except Exception:
            pass

        try:
            if ICON_PNG.exists():
                from PIL import Image, ImageTk
                self._icon_photo = ImageTk.PhotoImage(Image.open(ICON_PNG))
                self.iconphoto(True, self._icon_photo)
        except Exception:
            pass

    def _build_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        header = ctk.CTkFrame(self, corner_radius=18)
        header.grid(row=0, column=0, sticky="ew", padx=14, pady=(14, 8))
        header.grid_columnconfigure(0, weight=1)

        title = ctk.CTkLabel(
            header,
            text=APP_TITLE,
            font=ctk.CTkFont(size=22, weight="bold"),
        )
        title.grid(row=0, column=0, sticky="w", padx=18, pady=(16, 4))

        subtitle = ctk.CTkLabel(
            header,
            text="Φόρτωσε PDF διαταγής και Excel δύναμης, βρες τα κοινά πρόσωπα και κάνε εξαγωγή σε Word.",
            text_color=("gray25", "gray75"),
            font=ctk.CTkFont(size=12),
        )
        subtitle.grid(row=1, column=0, sticky="w", padx=18, pady=(0, 16))

        body = ctk.CTkFrame(self, corner_radius=18)
        body.grid(row=1, column=0, sticky="nsew", padx=14, pady=(0, 14))
        body.grid_columnconfigure(0, weight=0)
        body.grid_columnconfigure(1, weight=1)
        body.grid_rowconfigure(0, weight=1)

        self._build_sidebar(body)
        self._build_main_panel(body)

    def _build_sidebar(self, parent):
        sidebar = ctk.CTkFrame(parent, width=335, corner_radius=18)
        sidebar.grid(row=0, column=0, sticky="nsw", padx=(12, 10), pady=12)
        sidebar.grid_propagate(False)
        sidebar.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(sidebar, text="Ενέργειες", font=ctk.CTkFont(size=18, weight="bold")).grid(
            row=0, column=0, sticky="w", padx=16, pady=(18, 14)
        )

        self.btn_pdf = ctk.CTkButton(
            sidebar,
            text="1. Φόρτωση PDF Διαταγής",
            command=self.load_pdf,
            height=42,
            fg_color=THEME_BLUE,
        )
        self.btn_pdf.grid(row=1, column=0, sticky="ew", padx=16, pady=(0, 10))

        self.lbl_pdf = ctk.CTkLabel(sidebar, text="Δεν έχει επιλεγεί PDF", justify="left", wraplength=285)
        self.lbl_pdf.grid(row=2, column=0, sticky="w", padx=16, pady=(0, 12))

        self.btn_excel = ctk.CTkButton(
            sidebar,
            text="2. Φόρτωση Excel Δύναμης",
            command=self.load_excel,
            height=42,
            fg_color=THEME_BLUE,
        )
        self.btn_excel.grid(row=3, column=0, sticky="ew", padx=16, pady=(0, 10))

        self.lbl_excel = ctk.CTkLabel(sidebar, text="Δεν έχει επιλεγεί Excel", justify="left", wraplength=285)
        self.lbl_excel.grid(row=4, column=0, sticky="w", padx=16, pady=(0, 12))

        self.btn_run = ctk.CTkButton(
            sidebar,
            text="3. Εύρεση Κοινών",
            command=self.run_match,
            height=46,
            font=ctk.CTkFont(size=15, weight="bold"),
            fg_color="#2e8b57",
            hover_color="#28794b",
        )
        self.btn_run.grid(row=5, column=0, sticky="ew", padx=16, pady=(8, 12))

        self.btn_export_word = ctk.CTkButton(
            sidebar,
            text="4. Εξαγωγή σε Word",
            command=self.export_word,
            height=42,
            state="disabled",
        )
        self.btn_export_word.grid(row=6, column=0, sticky="ew", padx=16, pady=(0, 8))

        self.btn_export_excel = ctk.CTkButton(
            sidebar,
            text="5. Εξαγωγή και σε Excel",
            command=self.export_excel,
            height=42,
            state="disabled",
        )
        self.btn_export_excel.grid(row=7, column=0, sticky="ew", padx=16, pady=(0, 16))

        summary_box = ctk.CTkFrame(sidebar, corner_radius=14)
        summary_box.grid(row=8, column=0, sticky="ew", padx=16, pady=(4, 12))
        summary_box.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(summary_box, text="Σύνοψη", font=ctk.CTkFont(size=16, weight="bold")).grid(
            row=0, column=0, sticky="w", padx=14, pady=(12, 6)
        )
        self.summary_text = ctk.CTkTextbox(summary_box, height=235, wrap="word")
        self.summary_text.grid(row=1, column=0, sticky="ew", padx=12, pady=(0, 12))
        self.summary_text.insert("1.0", "Περίμενω αρχεία για επεξεργασία...")
        self.summary_text.configure(state="disabled")

    def _build_main_panel(self, parent):
        panel = ctk.CTkFrame(parent, corner_radius=18)
        panel.grid(row=0, column=1, sticky="nsew", padx=(0, 12), pady=12)
        panel.grid_columnconfigure(0, weight=1)
        panel.grid_rowconfigure(1, weight=1)

        topbar = ctk.CTkFrame(panel, fg_color="transparent")
        topbar.grid(row=0, column=0, sticky="ew", padx=14, pady=(14, 8))
        topbar.grid_columnconfigure(0, weight=1)

        self.status_label = ctk.CTkLabel(
            topbar,
            text="Έτοιμο για φόρτωση αρχείων.",
            font=ctk.CTkFont(size=13, weight="bold"),
        )
        self.status_label.grid(row=0, column=0, sticky="w")

        self.tabview = ctk.CTkTabview(panel, corner_radius=16)
        self.tabview.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 12))

        self.tab_results = self.tabview.add("Κοινοί")
        self.tab_only_pdf = self.tabview.add("Μόνο στο PDF")
        self.tab_only_excel = self.tabview.add("Μόνο στο Excel")
        self.tab_logs = self.tabview.add("Καταγραφές")

        self._build_results_tab(self.tab_results, is_main=True)
        self._build_only_pdf_tab(self.tab_only_pdf)
        self._build_only_excel_tab(self.tab_only_excel)
        self._build_logs_tab(self.tab_logs)

    def _build_results_tab(self, tab, is_main=False):
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(0, weight=1)
        frame = ctk.CTkFrame(tab, corner_radius=12)
        frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_rowconfigure(0, weight=1)

        style = ttk.Style()
        style.theme_use("default")
        style.configure("Treeview", rowheight=28, font=("Arial", 10))
        style.configure("Treeview.Heading", font=("Arial", 10, "bold"))

        table = ResultsTable(frame, TABLE_COLUMNS)
        table.grid(row=0, column=0, sticky="nsew", padx=(12, 0), pady=12)
        ybar = ttk.Scrollbar(frame, orient="vertical", command=table.yview)
        ybar.grid(row=0, column=1, sticky="ns", pady=12)
        xbar = ttk.Scrollbar(frame, orient="horizontal", command=table.xview)
        xbar.grid(row=1, column=0, sticky="ew", padx=(12, 0), pady=(0, 12))
        table.configure(yscrollcommand=ybar.set, xscrollcommand=xbar.set)

        if is_main:
            self.results_table = table

    def _build_only_pdf_tab(self, tab):
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(0, weight=1)
        self.only_pdf_text = ctk.CTkTextbox(tab, wrap="none")
        self.only_pdf_text.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)

    def _build_only_excel_tab(self, tab):
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(0, weight=1)
        self.only_excel_text = ctk.CTkTextbox(tab, wrap="none")
        self.only_excel_text.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)

    def _build_logs_tab(self, tab):
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(0, weight=1)
        self.log_text = ctk.CTkTextbox(tab, wrap="word")
        self.log_text.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        self.log("Η εφαρμογή ξεκίνησε.")

    def set_status(self, text: str):
        self.status_label.configure(text=text)
        self.update_idletasks()

    def log(self, text: str):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert("end", f"[{timestamp}] {text}\n")
        self.log_text.see("end")

    def set_summary(self, text: str):
        self.summary_text.configure(state="normal")
        self.summary_text.delete("1.0", "end")
        self.summary_text.insert("1.0", text)
        self.summary_text.configure(state="disabled")

    def load_pdf(self):
        file_path = filedialog.askopenfilename(
            title="Επιλογή PDF Διαταγής",
            filetypes=[("PDF files", "*.pdf")],
        )
        if not file_path:
            return
        self.pdf_path = file_path
        self.lbl_pdf.configure(text=file_path)
        self.log(f"Επιλέχθηκε PDF: {file_path}")
        self.set_status("Το PDF φορτώθηκε στη λίστα αναμονής.")

    def load_excel(self):
        file_path = filedialog.askopenfilename(
            title="Επιλογή Excel Δύναμης Υπηρεσίας",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
        )
        if not file_path:
            return
        self.excel_path = file_path
        self.lbl_excel.configure(text=file_path)
        self.log(f"Επιλέχθηκε Excel: {file_path}")
        self.set_status("Το Excel φορτώθηκε στη λίστα αναμονής.")

    def run_match(self):
        if not self.pdf_path or not self.excel_path:
            messagebox.showwarning("Λείπουν αρχεία", "Χρειάζεται και PDF και Excel πριν εκτελέσεις τη σύγκριση.")
            return

        try:
            self.set_status("Ανάγνωση PDF...")
            promotions_df = self.pdf_parser.parse(self.pdf_path)
            self.log(f"Διαβάστηκαν {len(promotions_df)} εγγραφές από το PDF.")

            self.set_status("Ανάγνωση Excel...")
            service_df = self.excel_loader.load(self.excel_path)
            self.log(f"Διαβάστηκαν {len(service_df)} εγγραφές από το Excel.")

            self.set_status("Σύγκριση αρχείων...")
            common_df, bundle = self.matcher.match(promotions_df, service_df)
            self.results_bundle = bundle

            self.results_table.set_rows(common_df)
            self._fill_text_tab(self.only_pdf_text, bundle["only_promotions"], from_pdf=True)
            self._fill_text_tab(self.only_excel_text, bundle["only_service"], from_pdf=False)
            self._refresh_summary(bundle["summary"])

            self.btn_export_word.configure(state="normal")
            self.btn_export_excel.configure(state="normal")
            self.set_status(f"Ολοκληρώθηκε. Βρέθηκαν {bundle['summary']['common_total']} κοινά άτομα.")
            self.log("Η σύγκριση ολοκληρώθηκε επιτυχώς.")
        except Exception as exc:  # noqa: BLE001
            self.log(f"Σφάλμα: {exc}")
            self.log(traceback.format_exc())
            self.set_status("Προέκυψε σφάλμα κατά τη σύγκριση.")
            messagebox.showerror("Σφάλμα", str(exc))

    def _refresh_summary(self, summary: dict):
        text = (
            f"Σύνολο εγγραφών PDF: {summary['promotions_total']}\n"
            f"Σύνολο εγγραφών Excel: {summary['service_total']}\n"
            f"Κοινά άτομα: {summary['common_total']}\n"
            f"Ταυτίσεις με Μητρώο: {summary['registry_matches']}\n"
            f"Ταυτίσεις με Ονοματεπώνυμο + Πατρώνυμο: {summary['name_matches']}\n"
            f"Μόνο στο PDF: {summary['only_promotions_total']}\n"
            f"Μόνο στο Excel: {summary['only_service_total']}"
        )
        self.set_summary(text)

    def _fill_text_tab(self, widget: ctk.CTkTextbox, df: pd.DataFrame, *, from_pdf: bool):
        widget.delete("1.0", "end")
        if df.empty:
            widget.insert("1.0", "Δεν υπάρχουν εγγραφές σε αυτή την κατηγορία.")
            return

        lines = []
        if from_pdf:
            ordered = df.sort_values(by=["pdf_row_no", "surname", "name"], na_position="last")
            for _, row in ordered.iterrows():
                lines.append(
                    f"Α/Α {row.get('pdf_row_no', '')} | Μητρώο {row.get('registry', '')} | "
                    f"{row.get('surname', '')} {row.get('name', '')} {row.get('patronymic', '')}"
                )
        else:
            ordered = df.sort_values(by=["surname", "name", "registry"], na_position="last")
            for _, row in ordered.iterrows():
                lines.append(
                    f"Μητρώο {row.get('registry', '')} | {row.get('surname', '')} {row.get('name', '')} {row.get('patronymic', '')} | "
                    f"{row.get('service_unit', '')}"
                )
        widget.insert("1.0", "\n".join(lines))

    def export_word(self):
        if not self.results_bundle:
            messagebox.showwarning("Δεν υπάρχουν αποτελέσματα", "Τρέξε πρώτα τη σύγκριση.")
            return

        filename = f"apotelesmata_sygkrisis_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
        output_path = filedialog.asksaveasfilename(
            title="Αποθήκευση Word",
            initialfile=filename,
            defaultextension=".docx",
            filetypes=[("Word Document", "*.docx")],
        )
        if not output_path:
            return

        try:
            self.word_exporter.export(
                output_path,
                self.results_bundle["common"],
                self.results_bundle["summary"],
                self.pdf_path or "",
                self.excel_path or "",
            )
            self.log(f"Έγινε εξαγωγή Word: {output_path}")
            self.set_status("Η εξαγωγή Word ολοκληρώθηκε.")
            messagebox.showinfo("Ολοκληρώθηκε", f"Το Word αποθηκεύτηκε εδώ:\n{output_path}")
        except Exception as exc:  # noqa: BLE001
            self.log(f"Σφάλμα εξαγωγής Word: {exc}")
            messagebox.showerror("Σφάλμα εξαγωγής", str(exc))

    def export_excel(self):
        if not self.results_bundle:
            messagebox.showwarning("Δεν υπάρχουν αποτελέσματα", "Τρέξε πρώτα τη σύγκριση.")
            return

        filename = f"apotelesmata_sygkrisis_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        output_path = filedialog.asksaveasfilename(
            title="Αποθήκευση Excel",
            initialfile=filename,
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx")],
        )
        if not output_path:
            return

        try:
            with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
                self.results_bundle["common"].to_excel(writer, sheet_name="Κοινοί", index=False)
                self.results_bundle["only_promotions"].to_excel(writer, sheet_name="Μόνο_PDF", index=False)
                self.results_bundle["only_service"].to_excel(writer, sheet_name="Μόνο_Excel", index=False)
                pd.DataFrame([self.results_bundle["summary"]]).to_excel(writer, sheet_name="Σύνοψη", index=False)
            self.log(f"Έγινε εξαγωγή Excel: {output_path}")
            self.set_status("Η εξαγωγή Excel ολοκληρώθηκε.")
            messagebox.showinfo("Ολοκληρώθηκε", f"Το Excel αποθηκεύτηκε εδώ:\n{output_path}")
        except Exception as exc:  # noqa: BLE001
            self.log(f"Σφάλμα εξαγωγής Excel: {exc}")
            messagebox.showerror("Σφάλμα εξαγωγής", str(exc))


def main():
    ctk.set_widget_scaling(1.0)
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
