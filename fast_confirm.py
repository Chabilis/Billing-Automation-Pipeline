# fast_confirm.py
# Lorkwen Trucking — 5-Second Private Confirm GUI
# Marlou only — NO browser, NO delay
# Double-click → drop PDF → fix 2 fields → Enter → DONE

import tkinter as tk
from tkinter import simpledialog, messagebox, filedialog
from pathlib import Path
from main_extractor import pdfText, extractdata, enhanceInfo

class EditableConfirm:
    def __init__(self, data):
        self.data = data
        self.root = tk.Tk()
        self.root.title("Lorkwen Trucking — 5-Second Confirm")
        self.root.geometry("580x620")
        self.root.configure(bg="#1e1e1e")
        self.root.resizable(False, False)

        # Title
        tk.Label(self.root, text="TRIP TICKET CONFIRMATION", font=("Consolas", 16, "bold"),
                 fg="#00ff00", bg="#1e1e1e").pack(pady=10)

        # Fields with double-click to edit
        fields = [
            ("Trip Ticket", data.get("trip_ticket", "")),
            ("Delivery Date", data.get("delivery_date", "")),
            ("Origin", data.get("origin", "")),
            ("Total Blocks", str(data.get("total_blocks", 0))),
            ("Reference Nos", data.get("ref_nos", "")),
            ("Seal Nos", ", ".join(data.get("seal_nos", [])) or "None"),
            ("Plate No", data.get("plate_no", "")),
            ("Driver", data.get("driver", "")),
            ("Helper 1", data.get("helper1", "")),
            ("Helper 2", data.get("helper2", "")),
            ("Shipper", data.get("shipper_full", "")),
            ("Route", f"{data.get('from_location','')} → {data.get('to_location','')}".strip(" →")),
        ]

        self.labels = {}
        for label_text, value in fields:
            frame = tk.Frame(self.root, bg="#1e1e1e")
            frame.pack(fill="x", padx=20, pady=4)

            tk.Label(frame, text=f"{label_text:<15}:", fg="#00ffff", bg="#1e1e1e",
                     font=("Consolas", 11), anchor="w").pack(side="left")

            val_label = tk.Label(frame, text=value or "[empty]", fg="#ffffff", bg="#2d2d2d",
                                 font=("Consolas", 11, "bold"), anchor="w", width=40, relief="sunken", padx=10)
            val_label.pack(side="left", fill="x", expand=True)

            # Double-click to edit
            def make_edit(val_label=val_label, key=label_text):
                def edit(event):
                    new_val = simpledialog.askstring(f"Edit {key}", f"Current: {val_label.cget('text')}\n\nEnter new value:",
                                                   initialvalue=val_label.cget('text'))
                    if new_val is not None:
                        val_label.config(text=new_val if new_val else "[empty]")
                        self.data[key] = new_val.strip() if new_val else ""
                return edit

            val_label.bind("<Double-1>", make_edit())

            self.labels[label_text] = val_label

        # Confirm button
        tk.Button(self.root, text="Generate Excel File", font=("Consolas", 14, "bold"),
                  bg="#00ff00", fg="black", height=2,
                  command=self.on_confirm).pack(pady=20)

        self.root.mainloop()

    def on_confirm(self):
        from excel_generator import generate_lts_excel
        
        final_data = {}
        for key, label in self.labels.items():
            text = label.cget("text")
            final_data[key] = "" if text in ("[empty]", "") else text
        
        file_path = generate_lts_excel(final_data)
        
        messagebox.showinfo(
            "EMPIRE COMPLETE",
            f"PERFECT FILE GENERATED!\n\n{file_path.name}\n\nOpen output/ folder → PRINT → DONE\n\nYou just killed manual work forever.",
            icon="info"
        )
        self.root.quit()

# === RUN IT ===
if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
    root.destroy()

    if file_path:
        raw = pdfText(file_path)
        ocr = extractdata(raw)
        info = enhanceInfo(ocr) or {}
        full_data = {**ocr, **info}
        full_data.update({
            "delivery_date": ocr.get("delivery_date", ""),
            "plate_no": info.get("plate_no", ""),
            "driver": info.get("driver", ""),
            "helper1": info.get("helper1", ""),
            "helper2": info.get("helper2", ""),
            "shipper_full": info.get("shipper_full", ""),
            "from_location": info.get("from_location", ""),
            "to_location": info.get("to_location", ""),
        })
        EditableConfirm(full_data)