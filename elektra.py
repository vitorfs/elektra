import sys

from nltk import word_tokenize

from utils import *

BASE_PATH = '/Users/vitorfs/Development/unioulu/elektra/data/'
TITLE_INDEX = 3
ABSTRACT_INDEX = 5

inclusion_criteria = read_criteria_from_file('%scriteria.txt' % BASE_PATH)
dataset = parse_spreadsheet('%sinput.xlsx' % BASE_PATH, TITLE_INDEX, ABSTRACT_INDEX)

criteria_list = list()

print('CRITERIA TOKENS:')

i = 0

for criteria in inclusion_criteria:
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
total_items_to_process = len(dataset) * len(criteria_list)
results = list()
for paper in dataset:
    i += 1
    abstract = normalize(paper[1])
    abs_tokens = word_tokenize(abstract)
    abs_filtered = remove_stopwords(abs_tokens)
    abs_stemmed = stemming(abs_filtered)
    document = set(abs_stemmed)

    title = ' '.join(paper[0].split())
    result_row = [i, paper[2], paper[3], title]

    sum_match_percentage = 0

    for criteria in criteria_list:
        current_progress += 1

        matches = criteria['size'] - len(criteria['keywords'].difference(document))
        match_percentage = matches / float(criteria['size'])
        sum_match_percentage += match_percentage

        result_row.append(matches)
        result_row.append(match_percentage)

        percent = (float(current_progress) / float(total_items_to_process)) * 100.0
        sys.stdout.write("\r%d%%" % percent)
        sys.stdout.flush()


    result_row.append(sum_match_percentage / len(criteria_list))
    results.append(result_row)

write_output(results, criteria_list)
