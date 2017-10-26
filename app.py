import subprocess, sys

from tkinter import ttk
import tkinter as tk
from tkinter import filedialog
from tkinter.scrolledtext import ScrolledText

from nltk import word_tokenize

import xlrd

from utils import *


class InclusionCriteriaDialog:
    def __init__(self, parent):
        top = self.top = tk.Toplevel(parent)
        self.parent = parent

        tk.Label(top, text='Inclusion Criteria:').pack()
        self.textarea = ScrolledText(top, wrap=tk.WORD)
        initial_text = '\n'.join(self.parent.criteria_list)
        self.textarea.insert(tk.END, initial_text)
        self.textarea.pack(padx=5, fill=tk.BOTH, expand=True)
        save_button = tk.Button(top, text='Save', command=self.save)
        save_button.pack(pady=5)

    def save(self):
        criteria = self.textarea.get('1.0', tk.END)
        data = criteria.splitlines()
        for i in range(len(data)):
            data[i] = ' '.join(data[i].split())
        self.parent.criteria_list = data
        self.top.destroy()


class Elektra(tk.Frame):
    def __init__(self, master=None, treeview=None):
        super().__init__(master)
        self.pack(side=tk.TOP, fill=tk.X)
        self.create_widgets()

        self.treeview = treeview
        self.criteria_list = list()

    def create_widgets(self):
        self.file_button = tk.Button(self)
        self.file_button['text'] = 'Select file'
        self.file_button['command'] = self.open_file_dialog
        self.file_button.pack(side=tk.LEFT)

        self.inclusion_criteria_button = tk.Button(self)
        self.inclusion_criteria_button['text'] = 'Inclusion criteria (0)'
        self.inclusion_criteria_button['command'] = self.inclusion_criteria_command
        self.inclusion_criteria_button.pack(side=tk.LEFT)

        self.execute_button = tk.Button(self)
        self.execute_button['text'] = 'Execute'
        self.execute_button['command'] = self.execute
        self.execute_button.pack(side=tk.LEFT)

    def inclusion_criteria_command(self):
        dialog = InclusionCriteriaDialog(self)
        self.master.wait_window(dialog.top)
        count = len(self.criteria_list)
        self.inclusion_criteria_button['text'] = 'Inclusion criteria (%s)' % count

    def open_file_dialog(self):
        _file = filedialog.askopenfile(mode='rb')
        if _file is not None:
            self.dataset = parse_excel(_file)
            for row in self.dataset:
                self.treeview.insert('', 'end', text=row[0], values=(row[1],))

    def execute(self):
        criteria_list = list()

        print('CRITERIA TOKENS:')

        i = 0

        for criteria in self.criteria_list:
            i += 1
            normalized = normalize(criteria)
            tokens = word_tokenize(normalized)
            filtered = remove_stopwords(tokens)
            stemmed = stemming(filtered)
            keywords = set(stemmed)

            criteria_info = {'id': i, 'keywords': keywords, 'size': len(keywords)}
            criteria_list.append(criteria_info)

            print('CRITERIA_%s:' % criteria_info['id'])
            print('LEN: (%s)' % criteria_info['size'])
            print('KEYWORDS: (%s)' % str(criteria_info['keywords']))

        print('-------------------------------------------------------')

        i = 0
        current_progress = 0
        total_items_to_process = len(self.dataset) * len(criteria_list)
        results = list()
        for paper in self.dataset:
            i += 1
            abstract = normalize(paper[1])
            abs_tokens = word_tokenize(abstract)
            abs_filtered = remove_stopwords(abs_tokens)
            abs_stemmed = stemming(abs_filtered)
            document = set(abs_stemmed)

            title = ' '.join(paper[0].split())
            result_row = [i, title,]

            sum_match_percentage = 0

            for criteria in criteria_list:
                current_progress += 1

                matches = criteria['size'] - len(criteria['keywords'].difference(document))
                match_percentage = matches / float(criteria['size'])
                sum_match_percentage += match_percentage

                result_row.append(matches)
                result_row.append(match_percentage)

                percent = (float(current_progress) / float(total_items_to_process)) * 100.0
                self.execute_button['text'] = 'Execute ({}%)'.format(percent)

            result_row.append(sum_match_percentage / len(criteria_list))
            results.append(result_row)

        write_output(results, criteria_list)
        opener = 'open' if sys.platform == 'darwin' else 'xdg-open'
        subprocess.call([opener, 'output.xls'])


def create_table(parent):
    tv = ttk.Treeview(parent)
    tv['columns'] = ('abstract',)

    tv.heading('#0', text='Titles', anchor='w')
    tv.column('#0', anchor='w')

    tv.heading('abstract', text='Abstract')
    tv.column('abstract', anchor='center', width=100)

    tv.grid(sticky = (tk.N, tk.S, tk.W, tk.E))

    parent.treeview = tv
    parent.grid_rowconfigure(0, weight=1)
    parent.grid_columnconfigure(0, weight=1)
    return tv


root = tk.Tk()
root.title('Systematic Literature Review Screening Tool')
root.geometry('800x600')

mainframe = tk.Frame(root, bg='#eee')
treeview = create_table(mainframe)

mainframe.pack(side=tk.BOTTOM, fill=tk.BOTH, expand=True)


app = Elektra(master=root, treeview=treeview)
app.mainloop()