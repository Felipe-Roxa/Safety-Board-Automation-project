import csv
import xlsxwriter

def inspection_data_mining_fsi():
    """This function generates a excel file"""

    inspection_data_mining_fsi_csv = 'C:/Users/felsique/Desktop/Safety Board/Input/Arquivos FSI/Inspections Data Mining FSI.csv'

    #Work on "Inspections Data Mining FSI.csv"
    with open(inspection_data_mining_fsi_csv) as f:
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

    assignee = [] #7
    location = [] #11

    for register in list_registers:
        assignee.append(register[7])
        location.append(register[11])
    # print(assignee)
    # print(location)

    list_occurance_assignee = set(assignee)
    list_occurance_location = set(location)
    # print(list_occurance_assignee)
    # print(list_occurance_location)

    # Creates a dictionary for each assignee and count how many FSI's were made bt each one:
    list_count_ocurance_assignee = []

    for occurance_assignee in list_occurance_assignee:
        assignee_dict = {'assignee_name': f'{occurance_assignee}', 'count':''}
        occurance_assignee_count = 0
        for register in list_registers:
            if register[7] == occurance_assignee:
                occurance_assignee_count += 1
                assignee_dict['count'] = occurance_assignee_count
                if assignee_dict not in list_count_ocurance_assignee:
                    list_count_ocurance_assignee.append(assignee_dict)
                else:
                    continue
            else:
                continue
            
    list_count_ocurance_assignee = sorted(list_count_ocurance_assignee, key = lambda item: item['count'], reverse = True)
    # print(list_count_ocurance_assignee)

    # Creates a list of dictionaries of each location, count it occurances, actions made, and percentage of non conformities:
    list_count_ocurance_and_actions_by_location = []

    for occurance_location in list_occurance_location:
        location_dict = {'location_name': f'{occurance_location}', 'count':'', 'total_actions':'', 'percentage':''}
        location_sum_total_actions = 0
        location_count = 0
        percentege = 0
        for register in list_registers:
            if register[11] == occurance_location:
                location_sum_total_actions += int(register[20])
                location_count += 1
                location_dict['total_actions'] = location_sum_total_actions
                location_dict['count'] = location_count
                if location_dict not in list_count_ocurance_and_actions_by_location:
                    list_count_ocurance_and_actions_by_location.append(location_dict)
                else:
                    continue
            else:
                continue
        
        for location_dict in list_count_ocurance_and_actions_by_location:
            questions = location_dict['count'] * 42
            location_dict['percentage'] = str("%.2f" % ((int(location_dict['total_actions']) / questions) * 100 )) + "%"

    list_count_ocurance_and_actions_by_location = sorted(list_count_ocurance_and_actions_by_location, key = lambda item: item['percentage'], reverse = True)

    # print(f"{list_count_ocurance_and_actions_by_location}\n")
    # for item in list_count_ocurance_assignee:
    #     print(f"{item}")
    # for item in list_count_ocurance_and_actions_by_location:
    #     print(f"{item}")

    # Start xlsxwriter library and export all previous data generated on this file to a excel
    workbook = xlsxwriter.Workbook('C:/Users/felsique/Desktop/Safety Board/Output/safety_board_fsi_inspections_data_mining.xlsx')

    worksheet01 = workbook.add_worksheet('Assignee count')
    row = 1
    col = 0
    worksheet01.write(0, 0, 'assignee_name')
    worksheet01.write(0, 1, 'count')
    for count_ocurance_assignee in list_count_ocurance_assignee:
        worksheet01.write(row, col, count_ocurance_assignee['assignee_name'])
        worksheet01.write(row, col + 1, count_ocurance_assignee['count'])
        row += 1
    worksheet01.write(0, 2, 'total_count')
    worksheet01.write(1, 2, len(list_registers))

    worksheet02 = workbook.add_worksheet('Location data')
    row = 1
    col = 0
    worksheet02.write(0, 0, 'location_name')
    worksheet02.write(0, 1, 'count')
    worksheet02.write(0, 2, 'total_actions')
    worksheet02.write(0, 3, 'percentage')
    for count_ocurance_and_actions_by_location in list_count_ocurance_and_actions_by_location:
        worksheet02.write(row, col, count_ocurance_and_actions_by_location['location_name'])
        worksheet02.write(row, col + 1, count_ocurance_and_actions_by_location['count'])
        worksheet02.write(row, col + 2, count_ocurance_and_actions_by_location['total_actions'])
        worksheet02.write(row, col + 3, count_ocurance_and_actions_by_location['percentage'])
        row += 1

    workbook.close()