import csv
import xlsxwriter
from encodings import utf_8

def question_level_fsi():
    """This function generates a excel file"""

    question_level_fsi_csv = 'C:/Users/felsique/Desktop/Safety Board/Input/Arquivos FSI/Question level FSI.csv'
    questions_categories_fsi_csv = 'C:/Users/felsique/Desktop/Safety Board/Input/Arquivos FSI/Questions categories FSI.csv'

    # Work on "Questions categories FSI.csv":
    with open(questions_categories_fsi_csv) as f:
        # print(f)
        reader = csv.reader(f, delimiter = ';')
        header_row = next(reader)

        list_registers = []
        for row in reader:
            if row != []:
                list_registers.append(row)
            else:
                continue

    list_questions_category = list_registers

    # for item in list_questions_category:
    #     print(f"{item}")
    #     print(f"{item[0]} AND {item[1]}")

    # Work on question_level_fsi_csv
    with open(question_level_fsi_csv, encoding='utf-16-le') as f:
        reader = csv.reader((line.replace('\0', '') for line in f), delimiter='\t')
        header_row = next(reader)

        list_registers = []
        for row in reader:
            if row != []:
                list_registers.append(row)
            else:
                continue
    # print(len(list_registers))
    # print(list_registers)
    # print(header_row)

    # Making a set of questions
    list_questions = []

    for register in list_registers:
        list_questions.append(register[5])

    set_question = set(list_questions)
    # for item in list_occurance_question:
    #     print(f"{item}")
    # print('')

    # Creates a dictionary of each type of question and count it
    list_count_and_category_ocurance_question = []

    for occurance_question in set_question:
        question_dict = {'question_name':f'{occurance_question}', 'count':'', 'category':''}
        occurance_question_count = 0
        for register in list_registers:
            if register[5] == occurance_question:
                occurance_question_count += 1
                question_dict['count'] = occurance_question_count
                if question_dict not in list_count_and_category_ocurance_question:
                    list_count_and_category_ocurance_question.append(question_dict)
                else:
                    continue
            else:
                continue
    # for item in list_count_and_category_ocurance_question:
    #     print(f"{item}")

    # Categorize questions from list_count_and_category_ocurance_question (first return)
    for count_and_category_ocurance_question in list_count_and_category_ocurance_question:
        for question_category in list_questions_category:
            if count_and_category_ocurance_question['question_name'] == question_category[1]:
                count_and_category_ocurance_question['category'] = question_category[0]
            else:
                continue
        
    list_count_and_category_ocurance_question = sorted(list_count_and_category_ocurance_question, key = lambda item: item['category'], reverse = True)
    # for item in list_count_and_category_ocurance_question:
    #     print(f"{item}")

    # Creates a list and a set of questions categories from list_questions_category first column
    list_categories = []
    for question_category in list_questions_category:
        list_categories.append(question_category[0])
    set_categories = set(list_categories)
    # print(list_categories)

    # Creates a list of dictionaries of each register and categorizes it:
    list_each_registers = []
    for register in list_registers:
        register_dict = {'ins_exe_id':f'{register[1]}', 'category':'', 'action_title':f'{register[10]}'}
        for question_category in list_questions_category:
            if register[5] == question_category[1]:
                register_dict['category'] = question_category[0]
                list_each_registers.append(register_dict)
            else:
                continue
    # for item in list_each_registers:
    #     print(item)
    # print(len(list_each_registers))

    # Creates a category dictionary and count how many questions are in each one (second return)
    list_count_category = []

    for category in set_categories:
        category_dict = {'category_name':f'{category}', 'count':''}
        question_category_count = 0
        for each_register in list_each_registers:
            if each_register['category'] == category:
                question_category_count += 1
                category_dict['count'] = question_category_count
                if category_dict not in list_count_category:
                    list_count_category.append(category_dict)
                else:
                    continue
            else:
                continue

    list_count_category = sorted(list_count_category, key = lambda item: item['count'], reverse = True)
    # for item in list_count_category:
    #     print(item)

    # Creates a set of dictionaries from list_each_registers to show all kinds of actions (caution: actions is a user input) (third return)
    list_actions = []

    for each_register in list_each_registers:
        list_actions.append(each_register['action_title'])
    list_set_actions = set(list_actions)

    # for item in set_actions:
    #     print(item)

    # # Lenght comparsion between a list and a set of actions (last one: 83 x 71)
    # set_actions = set(list_actions)
    # print(set_actions)
    # list2_actions = set_actions
    # print(len(list2_actions))

    # Start xlsxwriter library and export all previous data generated on this file to a excel
    workbook = xlsxwriter.Workbook('C:/Users/felsique/Desktop/Safety Board/Output/safety_board_fsi_question_level.xlsx')

    worksheet01 = workbook.add_worksheet('Questions count and category')

    row = 1
    col = 0
    worksheet01.write(0, 0, 'question_name')
    worksheet01.write(0, 1, 'count')
    worksheet01.write(0, 2, 'category')
    for count_and_category_ocurance_question in list_count_and_category_ocurance_question:
        worksheet01.write(row, col, count_and_category_ocurance_question['question_name'])
        worksheet01.write(row, col + 1, count_and_category_ocurance_question['count'])
        worksheet01.write(row, col + 2, count_and_category_ocurance_question['category'])
        row += 1
    # for count_and_category_ocurance_question in list_count_and_category_ocurance_question:
    #     print(count_and_category_ocurance_question)
    #     for key, value in count_and_category_ocurance_question.items():
    #         print(f"This is the key: {key} and its value: {value}")

    worksheet02 = workbook.add_worksheet('Category count')
    row = 1
    col = 0
    worksheet02.write(0, 0, 'category_name')
    worksheet02.write(0, 1, 'count')
    for category_dict in list_count_category:
        worksheet02.write(row, col, category_dict['category_name'])
        worksheet02.write(row, col + 1, category_dict['count'])
        row += 1

    worksheet03 = workbook.add_worksheet('Actions list')
    row = 1
    col = 0
    worksheet03.write(0, 0, 'action_title')
    for each_action in list_set_actions:
        worksheet03.write(row, col, each_action)
        row += 1

    workbook.close()