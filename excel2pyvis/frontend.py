import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import pandas as pd
import networkx as nx
import matplotlib.pyplot as plt



def visualize_data(self):
    source_column = self.source_multiselect.get()
    target_column = self.target_multiselect.get()
    if source_column and target_column:
        graph = nx.from_pandas_edgelist(df, source=source_column, target=target_column)
        plt.figure(figsize=(10, 6))
        pos = nx.spring_layout(graph)
        nx.draw_networkx(graph, pos, with_labels=True, node_color='lightblue', node_size=500, edge_color='gray')
        plt.title("Network Visualization")
        plt.show()

class MyApp(tk.Tk):
    def __init__(self):
        self.root = tk.Tk.__init__(self)
        self.title("Excel2Pyvis")
        self.geometry("800x600")
        self.resizable(False, False)
        self.create_widgets()


    def create_widgets(self):
        self.frame = ttk.Frame(self.root)
        self.frame.pack()

        self.widgets_frame = ttk.LabelFrame(self.frame, text="Source and Target Nodes")
        self.widgets_frame.grid(row=0, column=0, padx=20, pady=10)

        self.import_button = ttk.Button(self.widgets_frame, text="Import File", command=self.import_file)
        self.import_button.grid(row=1, column=0, padx=20, pady=10)

        self.source_label = ttk.Label(self.widgets_frame, text="")
        self.source_label.grid(row=2,column=0, padx=20, pady=10)



        self.source_multiselect = tk.Listbox(self.widgets_frame, selectmode=tk.MULTIPLE,exportselection=0) 
        self.source_multiselect.grid(row=3,column=0, padx=20, pady=10)

        self.target_label = ttk.Label(self.widgets_frame, text="")
        self.target_label.grid(row=4,column=0, padx=20, pady=10)

        self.target_multiselect = tk.Listbox(self.widgets_frame, selectmode=tk.MULTIPLE,exportselection=0)
        self.target_multiselect.grid(row=5,column=0, padx=20, pady=10)

        self.visualize_button = tk.Button(self.widgets_frame, text="Visualize Data", command=visualize_data)
        self.visualize_button.grid(row=6,column=0, padx=20, pady=10)

        self.html_label = ttk.Label(self.widgets_frame, text="")
        self.html_label.grid(row=8,column=0, padx=20, pady=10)

    def import_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx"), ("CSV Files", "*.csv")])
        if file_path:
            self.df = pd.read_excel(file_path) if file_path.endswith('.xlsx') else pd.read_csv(file_path)
            columns = self.df.columns.tolist()
            self.source_label.config(text="Select Source Column:")
            self.target_label.config(text="Target Column, if you want the source to target all remaining columns leave blank:")
            self.source_multiselect.delete(0, tk.END)
            self.target_multiselect.delete(0, tk.END)
            for value in columns:
                self.source_multiselect.insert(tk.END, value)
            for value in columns:
                self.target_multiselect.insert(tk.END, value)
            self.html_label.config(text="HTML file will be saved in the same directory as the imported file")

    def visualize_data(self):
        source_column = self.source_multiselect.get()
        target_column = self.target_multiselect.get()
        if source_column and target_column:
            graph = nx.from_pandas_edgelist(self.df, source=source_column, target=target_column)
            plt.figure(figsize=(10, 6))
            pos = nx.spring_layout(graph)
            nx.draw_networkx(graph, pos, with_labels=True, node_color='lightblue', node_size=500, edge_color='gray')
            save_local = filedialog.asksaveasfilename(defaultextension=".html", filetypes=[("HTML Files", "*.html")])
            nx.savegraph_html(graph, save_local)
            plt.title("Network Visualization")
            plt.show()



if __name__ == "__main__":
    app = MyApp()
    app.mainloop()