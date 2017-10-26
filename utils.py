import xlrd
import xlwt
from nltk.corpus import stopwords
from nltk.stem.snowball import SnowballStemmer
import string


def parse_excel(_file):
    wb = xlrd.open_workbook(filename=None, file_contents=_file.read())
    ws = wb.sheets()[0]
    data = list()
    for row in range(ws.nrows):
        title = ws.cell(row, 3).value
        abstract = ws.cell(row, 5).value
        data.append([title, abstract,])
    return data[1:]


def parse_spreadsheet(filename, title_index, abstract_index, sheet_index=0):
    wb = xlrd.open_workbook(filename=filename)
    ws = wb.sheets()[sheet_index]
    data = list()
    for row in range(ws.nrows):
        title = ws.cell(row, title_index).value
        abstract = ws.cell(row, abstract_index).value
        level_1_screening = ws.cell(row, 0).value
        final_selection = ws.cell(row, 1).value
        data.append([title, abstract, level_1_screening, final_selection])
    return data[1:]


def read_criteria_from_file(filename):
    with open(filename, 'r') as f:
        lines = f.readlines()
    return lines


def write_output(results, criteria_list):
    criteria_headers = list()

    for criteria in criteria_list:
        criteria_headers.append('CRITERIA_%s_SUM' % criteria['id'])
        criteria_headers.append('CRITERIA_%s_SIMILARITY' % criteria['id'])

    header = ['ARTICLE_ID', 'TITLE',] + criteria_headers + ['TOTAL_SIMILARITY',]

    results.insert(0, header)

    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet('Screening')

    for i, row in enumerate(results):
        for j, col in enumerate(row):
            sheet.write(i, j, col)

    workbook.save('output.xls')


def normalize(text):
    translator = str.maketrans('', '', string.punctuation)
    text = text.translate(translator)
    text = text.lower()
    text = ' '.join(text.split())
    return text


def remove_stopwords(word_list):
    stops = set(stopwords.words('english'))
    filtered = [word for word in word_list if word not in stops]
    return filtered


def stemming(word_list):
    stemmer = SnowballStemmer('english')
    stemmed = map(lambda word: stemmer.stem(word), word_list)
    return list(stemmed)
