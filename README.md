import os
import cx_Oracle
import pandas as pd
import openpyxl
from tkinter import Tk, Label, Button, filedialog, messagebox

class PrimaveraExcelTool:
    def __init__(self, root):
        self.root = root
        self.root.title("Primavera-Excel Integration Tool")

        # GUI Components
        self.label = Label(root, text="Primavera-Excel Integration", font=("Helvetica", 16))
        self.label.pack(pady=20)

        self.import_btn = Button(root, text="Import from Primavera", command=self.import_from_primavera)
        self.import_btn.pack(pady=10)

        self.export_btn = Button(root, text="Export to Primavera", command=self.export_to_primavera)
        self.export_btn.pack(pady=10)

        self.exit_btn = Button(root, text="Exit", command=root.quit)
        self.exit_btn.pack(pady=20)

    def import_from_primavera(self):
        try:
            file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", 
                                                     filetypes=[("Excel Files", "*.xlsx")])
            if not file_path:
                return

            connection = cx_Oracle.connect("username/password@hostname:port/SID")
            cursor = connection.cursor()
            query = "SELECT * FROM PROJECTS"  # Replace with your Primavera query
            cursor.execute(query)
            data = cursor.fetchall()

            # Convert data to Excel
            columns = [col[0] for col in cursor.description]
            df = pd.DataFrame(data, columns=columns)
            df.to_excel(file_path, index=False)

            connection.close()
            messagebox.showinfo("Success", f"Data exported to {file_path}")

        except Exception as e:
            messagebox.showerror("Error", str(e))

    def export_to_primavera(self):
        try:
            file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
            if not file_path:
                return

            connection = cx_Oracle.connect("username/password@hostname:port/SID")
            cursor = connection.cursor()

            # Read data from Excel
            df = pd.read_excel(file_path)
            
            for _, row in df.iterrows():
                # Replace this with an appropriate Primavera SQL INSERT/UPDATE statement
                query = "INSERT INTO PROJECTS (COLUMN1, COLUMN2) VALUES (:1, :2)"
                cursor.execute(query, (row['COLUMN1'], row['COLUMN2']))

            connection.commit()
            connection.close()
            messagebox.showinfo("Success", "Data imported to Primavera successfully!")

        except Exception as e:
            messagebox.showerror("Error", str(e))

# Main function
if __name__ == "__main__":
    root = Tk()
    app = PrimaveraExcelTool(root)
    root.mainloop()
