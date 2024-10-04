import tkinter as tk
import pandas as pd
from tkinter import filedialog, messagebox
# from main import generate_gantt_chart, save_csv, extract_data_from_api
from op_efficiency import process_data_em, plot_employee_efficiency, process_data_op, plot_operation_efficiency
from gantt_chart import process_df,  generate_gantt_chart
from idle_time_report import create_summary, generate_mean_chart
from item_efficiency import item_graph_efficiency, save_to_csv


class MyApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Meiban Jobs Data")

        # Gantt Chart
        self.gantt_button = tk.Button(root, text="Generate Gantt Chart", command=self.upload_gantt_file)
        self.gantt_button.pack(pady=1)

        self.guideline_label = tk.Label(root, text="Please upload the Latest Job Activity File", bg="#0062A8",
                                        fg="#cfcfcf")
        self.guideline_label.pack(pady=5)

        # white space
        self.space_label = tk.Label(root, text="", height=1, bg="#0062A8")  # Creates an empty label with some height
        self.space_label.pack()

        # Generate Item Efficiency Chart
        self.csv_button = tk.Button(root, text="Generate Item Efficiency Chart",
                                    command=self.upload_item_efficiency_graph_file)
        self.csv_button.pack(pady=1)

        self.guideline_label = tk.Label(root, text="Please upload the Job Cost File", bg="#0062A8",
                                        fg="#cfcfcf")
        self.guideline_label.pack(pady=5)

        # white space
        self.space_label = tk.Label(root, text="", height=1, bg="#0062A8")  # Creates an empty label with some height
        self.space_label.pack()

        # Generate Item Efficiency Summary
        self.csv_button = tk.Button(root, text="Summarize Item Efficiency",
                                    command=self.save_item_efficiency_csv)
        self.csv_button.pack(pady=1)

        self.guideline_label = tk.Label(root, text="Please upload the Job Cost File", bg="#0062A8",
                                        fg="#cfcfcf")
        self.guideline_label.pack(pady=5)

        # white space
        self.space_label = tk.Label(root, text="", height=1, bg="#0062A8")  # Creates an empty label with some height
        self.space_label.pack()

        # Operation & Operator Efficiency
        self.upload_button = tk.Button(root, text="Operation & Operator Efficiency", command=self.upload_efficiency_file)
        self.upload_button.pack(pady=1)

        self.guideline_label = tk.Label(root, text="Please upload the Job Costs File", bg="#0062A8",
                                        fg="#cfcfcf")
        self.guideline_label.pack(pady=5)

        # white space
        self.space_label = tk.Label(root, text="", height=1, bg="#0062A8")  # Creates an empty label with some height
        self.space_label.pack()

        # Generate mean Idle Time Chart
        self.upload_button = tk.Button(root, text="Generate Mean Idle Time Chart", command=self.upload_idle_file_meanChart)
        self.upload_button.pack(pady=1)

        self.guideline_label = tk.Label(root, text="Please upload the Time Management File", bg="#0062A8",
                                        fg="#cfcfcf")
        self.guideline_label.pack(pady=5)

        # white space
        self.space_label = tk.Label(root, text="", height=1, bg="#0062A8")  # Creates an empty label with some height
        self.space_label.pack()

        # Summarize Idl Time
        self.csv_button = tk.Button(root, text="Summarize Idle Time", command=self.upload_idle_file_summary)
        self.csv_button.pack(pady=1)

        self.guideline_label = tk.Label(root, text="Please upload the Time Management File", bg="#0062A8",
                                        fg="#cfcfcf")
        self.guideline_label.pack(pady=5)

        self.loading_screen = None

    def show_loading_screen(self):
        if not hasattr(self, 'loading_screen') or self.loading_screen is None:
            self.loading_screen = tk.Toplevel(self.root)
            self.loading_screen.title("Loading...")
            self.loading_screen.geometry("200x100")
            label = tk.Label(self.loading_screen, text="Processing, please wait...", padx=20, pady=20)
            label.pack()
            self.root.update_idletasks()

    def hide_loading_screen(self):
        if hasattr(self, 'loading_screen') and self.loading_screen:
            self.loading_screen.destroy()
            self.loading_screen = None
            print("Loading screen hidden.")

    def upload_efficiency_file(self):
        try:
            file_path = filedialog.askopenfilename(
                filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
            )

            if file_path:
                # Read and process the Excel file
                df_em = pd.read_excel(file_path, sheet_name='Operator Time')
                df_op = pd.read_excel(file_path, sheet_name='Operations')

                # Process and plot data
                if df_em is not None:
                    df_em_processed = process_data_em(df_em)
                    plot_employee_efficiency(df_em_processed)

                if df_op is not None:
                    df_op_processed = process_data_op(df_op)
                    plot_operation_efficiency(df_op_processed)

        except Exception as e:
            messagebox.showinfo("Error", f"An error occurred: {str(e)}")

    def upload_gantt_file(self):
        try:
            file_path = filedialog.askopenfilename(
                filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
            )

            if file_path:
                # Read and process the Excel file
                df_job = pd.read_excel(file_path, sheet_name='Latest Job Activity')
                df_op = pd.read_excel(file_path, sheet_name='Open Operations')

                # Process and plot data
                if df_job is not None and df_op is not None:
                    df_processed = process_df(df_job,df_op)
                    generate_gantt_chart(df_processed)

        except Exception as e:
            messagebox.showinfo("Error", f"An error occurred: {str(e)}")

    def upload_item_efficiency_graph_file(self):
        try:
            file_path = filedialog.askopenfilename(
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
            )
            if file_path:
                # Read and process the Excel file
                df_item = pd.read_excel(file_path, sheet_name='Operations')

                if df_item is not None:
                    df_item_processed = item_graph_efficiency(df_item)
                    item_graph_efficiency(df_item_processed)

        except Exception as e:
            print(f"An error occurred: {str(e)}")

    def save_item_efficiency_csv(self):
        try:
            file_path = filedialog.askopenfilename(
                filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")])

            if file_path:
                # Read and process the Excel file
                df = pd.read_excel(file_path, sheet_name='Operations')

                # Process and plot data
                if df is not None:
                    excel_file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                                   filetypes=[("Excel files", "*.xlsx")])
                    df_processed = save_to_csv(df, excel_file_path)
                    save_to_csv(df_processed, excel_file_path)

        except Exception as e:
            messagebox.showinfo("Error", f"An error occurred: {str(e)}")

    def upload_idle_file_meanChart(self):
        try:
            file_path = filedialog.askopenfilename(
                filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
            )

            if file_path:
                # Read and process the Excel file
                df = pd.read_excel(file_path)

                # Process and plot data
                if df is not None:
                    df_processed = generate_mean_chart(df)
                    generate_mean_chart(df_processed)

        except Exception as e:
            messagebox.showinfo("Error", f"An error occurred: {str(e)}")

    def upload_idle_file_summary(self):
        try:
            file_path = filedialog.askopenfilename(
                filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")])

            if file_path:
                # Read and process the Excel file
                df = pd.read_excel(file_path)

                # Process and plot data
                if df is not None:
                    excel_file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                                   filetypes=[("Excel files", "*.xlsx")])
                    df_processed = create_summary(df, excel_file_path)
                    create_summary(df_processed, excel_file_path)

        except Exception as e:
            messagebox.showinfo("Error", f"An error occurred: {str(e)}")

    #def save_csv_file(self):
    #    csv_file_path = filedialog.asksaveasfilename(defaultextension=".csv",
    #                                                 filetypes=[("CSV files", "*.csv")])
    #    if csv_file_path:
    #        self.show_loading_screen()
    #        Thread(target=self.process_csv, args=(csv_file_path,)).start()

    #def process_csv(self, csv_file_path):
    #    try:
    #        save_csv(csv_file_path)
    #    finally:
    #        self.hide_loading_screen()

    #def show_gantt_chart(self):
    #    self.show_loading_screen()
    #    Thread(target=self.process_gantt_chart).start()

    #def process_gantt_chart(self):
    #    try:
    #        upload_gantt_file()
    #    finally:
    #        self.hide_loading_screen()

    #def extract_data(self):
    #    self.show_loading_screen()
    #    Thread(target=self.process_data_extraction).start()

    #def process_data_extraction(self):
    #    try:
    #        api_url = "https://api.example.com/data"
    #        df = extract_data_from_api(api_url)
    #        print("Data extracted:", df.head())
    #    finally:
    #        self.hide_loading_screen()

def setup_and_start_app():
    root = tk.Tk()
    height = 550
    width = 530
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry('{}x{}+{}+{}'.format(width, height, x, y))
    root.config(background="#0062a8")

    welcome_label = tk.Label(text="Meiban Jobs Data", bg="#0062a8", font=("Helvetica", 20, "bold"), fg="#FFFFFF")
    welcome_label.pack(pady=20)
    root.resizable(width=False, height=False)

    app = MyApp(root)

    root.mainloop()


if __name__ == "__main__":
    setup_and_start_app()