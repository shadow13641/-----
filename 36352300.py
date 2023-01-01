import pandas as pd
import tkinter as tk
from tkinter import ttk
import numpy
import pickle
import tkinter.messagebox as msg
import seaborn as sns

#part 1
dir = 'C:\\Users\\shado\\Downloads\\EnergyCallCentre.xlsx'
data=pd.read_excel(dir, engine="openpyxl")
data_sample = data.sample(n=100, random_state=2301)
h=data_sample.describe()
data.skew()
data_sample['Month'].value_counts().plot.pie()
data_sample['Month'].value_counts().plot.bar()
data_sample['VHT'].value_counts().plot.pie()
data_sample['ToD'].value_counts().plot.bar()
data_sample.groupby(['VHT'])['Agents'].mean()
sns.histplot(data=data_sample['Agents'], kde=True)
data_sample['Agents'].corr(data_sample['CallsOffered'])**2
sns.histplot(data=data_sample['CallsOffered'], kde=True)
data_sample.groupby(['VHT'])['CallsOffered'].mean()
data_sample['CallsOffered'].corr(data_sample['CallsAbandoned'])**2
data_sample['Agents'].corr(data_sample['CallsAbandoned'])**2
data_sample['CallsHandled'].corr(data_sample['CallsAbandoned'])**2
data_sample.groupby(['VHT'])['CallsAbandoned'].mean()
s=data_sample.corr()
data_sample.groupby(['VHT'])['ASA'].mean()

#part2
class Windows(tk.Tk):
    def __init__(self):
        super(Windows, self).__init__()
        self.tree = self.month = self.font = self.data = self.item = self.var = self.varChosen = self.min_value \
            = self.max_value = self.search = None
        self.init_shape()
        self.init_data()
        self.create_widget()

    def init_data(self):
        self.font = ('微软雅黑', 14)
        self.data = pd.read_excel('./1.xlsx')
        self.data.columns = ['Id'] + self.data.columns[1:].tolist()
        self.item = [i for i in self.data.dtypes.items() if i[1] != numpy.dtype('O')]
        self.pickfile = open('logging.pkl', 'wb')

    def init_shape(self):
        self.title('python')
        width = int(self.winfo_screenwidth() * 3 / 4)
        height = int(self.winfo_screenheight() * 3 / 4)
        space_l = int(self.winfo_screenwidth() * 1 / 8)
        space_t = int(self.winfo_screenheight() * 1 / 8)
        self.geometry(f'{width}x{height}+{space_l}+{space_t}')

    def create_widget(self):
        labelframe = ttk.Labelframe(self, text='condition', padding=10)
        labelframe.place(relx=0.1, rely=0.03)
        tip_date = ttk.Label(labelframe, text='Month:', font=self.font)
        tip_date.grid(row=0, column=0)
        self.month = ttk.Entry(labelframe)
        self.month.grid(row=0, column=1)
        ttk.Label(labelframe, text='varaible:', font=self.font).grid(row=0, column=2, padx=15)
        self.var = tk.StringVar()
        self.varChosen = ttk.Combobox(labelframe, width=15, textvariable=self.var)
        self.varChosen['values'] = [i[0] for i in self.item]
        self.varChosen.grid(row=0, column=3)
        self.varChosen.current(1)
        ttk.Label(labelframe, text='min number:', font=self.font).grid(row=0, column=4, padx=10)
        self.min_value = ttk.Entry(labelframe, width=10)
        self.min_value.grid(row=0, column=5)

        ttk.Label(labelframe, text='max number:', font=self.font).grid(row=0, column=6, padx=10)
        self.max_value = ttk.Entry(labelframe, width=10)
        self.max_value.grid(row=0, column=7)

        self.search = ttk.Button(labelframe, text='check', command=self.search_data)
        self.search.grid(row=0, column=8, padx=15)

        menu = tk.Menu(self)
        edit_menu = tk.Menu(menu, tearoff=False)
        edit_menu.add_command(label='Add', command=self.add_data)
        edit_menu.add_command(label='Modify', command=self.change_data)
        edit_menu.add_command(label='Delete', command=self.delete_data)
        menu.add_cascade(label='Edit', menu=edit_menu)

        self.configure(menu=menu)
        sb = tk.Scrollbar(self)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree = ttk.Treeview(self,
                                 show="headings",
                                 columns=[i for i in self.data.columns],
                                 height=28, yscrollcommand=sb.set)
        # set colum names
        for i in self.data.columns:
            self.tree.column(i, width=int(int(self.winfo_screenwidth() * 2 / 3) / len(self.data.columns)),
                             anchor='center')
            # Set the displayed name for the column name
            self.tree.heading(i, text=i)
        self.tree.place(relx=0.05, rely=0.15)
        sb.config(command=self.tree.yview)

    def search_data(self):
        self.clear_tree()
        domain = []
        pickle.dump(f'check record,check condition:{str(domain)}', self.pickfile)
        if self.month.get():
            domain.append("(self.data['Month']==self.month.get())")
        if self.min_value.get():
            domain.append("(self.data[self.var.get()]>=float(self.min_value.get()))")
        if self.max_value.get():
            domain.append("(self.data[self.var.get()]<=float(self.max_value.get()))")
        if not domain:
            self.show_all(self.data.values.tolist())
        else:
            df = self.data.loc[(eval('&'.join(domain))), :]
            self.show_all(df.values.tolist())

    def show_all(self, list):
        for index, line in enumerate(list):
            self.tree.insert('', index + 1,
                             values=line)

    def add_data(self):
        self.add = tk.Tk()
        self.add.title('Add record')
        self.add.geometry('400x400+400+200')
        a = {}

        def commit():
            value = {}
            for i, j in a.items():
                try:
                    value.update({i: eval(j.get())})
                except:
                    value.update({i: j.get()})
            self.data = self.data.append(value, ignore_index=True)
            pickle.dump('add a record', self.pickfile)
            self.add.destroy()

        for index, i in enumerate(self.data.columns):
            ttk.Label(self.add, text=i).grid(row=index, column=1, pady=5, padx=40)
            t = ttk.Entry(self.add)
            a.update({i: t})
            t.grid(row=index, column=2, pady=5)
        ttk.Button(self.add, text='Add', command=commit).grid(row=len(self.data.columns), column=1, columnspan=2)
        self.add.mainloop()

    def change_data(self):
        if not self.tree.selection():
            msg.showinfo('Hint', 'Please select a record first, then click Modify!')
        elif len(self.tree.selection()) > 1:
            msg.showinfo('Hint', 'Only one record can be modified at a time!')
        else:
            a = {}
            line = self.tree.item(self.tree.selection())['values']
            self.change = tk.Tk()
            self.change.title('Modify record')
            self.change.geometry('400x400+400+200')
            for index, i in enumerate(self.data.columns):
                ttk.Label(self.change, text=i).grid(row=index, column=1, pady=5, padx=40)
                t = ttk.Entry(self.change)
                a.update({i: t})
                t.grid(row=index, column=2, pady=5)
            for i, j in zip(a.values(), line):
                i.insert(0, str(j))

            condition = (self.data.Id == self.tree.item(self.tree.selection())['values'][0])

            def commit():
                value = {}
                for i, j in a.items():
                    try:
                        value.update({i: eval(j.get())})
                    except:
                        value.update({i: j.get()})
                for i, j in value.items():
                    self.data.loc[condition, i] = j
                self.clear_tree()
                self.show_all(self.data.values.tolist())
                pickle.dump('Modify record',self.pickfile)
                self.change.destroy()


            ttk.Button(self.change, text='Application', command=commit).grid(row=len(self.data.columns), column=1,
                                                                          columnspan=2)
            self.change.mainloop()

    def delete_data(self):
        if not self.tree.selection():
            msg.showinfo('Hint', 'Please select the record first, then click delete!')
        else:
            for item in self.tree.selection():
                a = self.tree.item(item)['values'][0]
                for i in self.data.values.tolist():
                    if i[0] == a:
                        self.tree.delete(item)
                        pickle.dump(f'delete record that id is {a}', self.pickfile)
                        self.data = self.data.drop(self.data[self.data['Id'] == a].index)

    def clear_tree(self):
        obj = self.tree.get_children()
        for o in obj:
            self.tree.delete(o)

    def __del__(self):
        self.pickfile.close()
        self.data.to_excel('1.xlsx',index=False)



if __name__ == '__main__':
    Windows().mainloop()

