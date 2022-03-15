from tkinter import *
from tkinter import ttk
from tkinter.filedialog import asksaveasfile, asksaveasfilename

import xlwt


class Form:
    def __init__(self, cursor):
        self.win = Tk()
        self.win.title("Table viewer")
        self.win.geometry("700x500")
        self.table = ttk.Treeview(self.win)
        self.cursor = cursor
        self.maxtablewidth = 500

        list = self.get_table_list()
        values = [row[1] for row in list]
        self.combo = ttk.Combobox(self.win, values=values, state="readonly")
        self.combo.bind("<<ComboboxSelected>>", self.table_changed)
        self.combo.current(0)
        self.button = Button(self.win, text="Save in table in excel", command=lambda: self.save_file())

        self.make_table(self.combo.get())

        self.pack()

    def pack(self):
        self.combo.pack(pady=30)
        self.table.pack(pady=50)
        self.button.pack(pady=20)

    def get_table_list(self):
        query = "SELECT table_schema, table_name " \
                "FROM information_schema.tables " \
                "WHERE (table_schema = 'public') " \
                "ORDER BY table_schema, table_name;"
        self.cursor.execute(query)
        return self.cursor.fetchall()

    def get_table_by_name(self, name):
        query = "SELECT *FROM "
        query += name
        self.cursor.execute(query)
        headers = self.cursor.description
        records = self.cursor.fetchall()
        return headers, records

    def clear_table(self):
        self.table.delete(*self.table.get_children())

    def table_changed(self, event):
        table_name = event.widget.get()
        self.make_table(table_name)

    def make_table(self, table_name):
        self.clear_table()
        h, r = self.get_table_by_name(table_name)
        headers = [header[0] for header in h]
        self.table['columns'] = headers
        self.table.column('#0', width=0, stretch=NO)
        size = len(headers)
        width = int(self.maxtablewidth / size)
        for i in range(size):
            i_str = "#" + str(i+1)
            self.table.heading(i_str, text=headers[i])
            self.table.column(i_str, minwidth=0, width=width, stretch=NO, anchor=CENTER)
        for index, row in enumerate(r):
            self.table.insert(parent='', index=index, iid=index, text='', values=row)
        self.table['show'] = 'headings'

    def save_file(self):
        f = asksaveasfilename(initialfile='Untitled.xls',
                              defaultextension=".xls", filetypes=[("All Files", "*.*"), ("Excel Documents", "*.xls")])
        workbook = xlwt.Workbook()
        sheetname = self.combo.get()
        worksheet = workbook.add_sheet(sheetname)

        headers = self.table['columns']
        rowcount = len(self.table.get_children())
        records = []
        for i in range(rowcount):
            records.append(self.table.item(i)['values'])
        # records = self.table['values']

        for colidx, heading in enumerate(headers):
            worksheet.write(0, colidx, heading)

        for rowidx, row in enumerate(records):
            for colindex, col in enumerate(row):
                worksheet.write(rowidx + 1, colindex, col)

        workbook.save(f)

    def loop(self):
        self.win.mainloop()
