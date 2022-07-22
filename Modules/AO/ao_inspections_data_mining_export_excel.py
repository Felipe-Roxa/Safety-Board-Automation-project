import csv
import xlsxwriter

def inspection_data_mining_ao(inspection_data_mining_ao_csv, zones_csv):
    """This function generates a excel file"""

    #Work on "Zones.csv":
    with open(zones_csv) as f:
        # print(f)
        reader = csv.reader(f, delimiter = ';')
        header_row = next(reader)

        list_registers = []
        for row in reader:
            if row != []:
                list_registers.append(row)
            else:
                continue

    list_zones = list_registers

    # for item in list_zones:
    #     print(f"{item}")
    #     print(f"{item[0]} AND {item[1]}")

    # Work on "Inspections Data Mining AO.csv"
    with open(inspection_data_mining_ao_csv) as f:
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

    # Creates a list of assignee total actions, sets zones for each one, and generate a percentage of nonconformities
    list_assignee_total_actions = []

    for register in list_registers:
        for zone in list_zones:
            if register[7] == zone[1]:
                questions = 36
                percentage = str("%.2f" % ((int(register[20]) / questions) * 100 )) + "%"
                assigne_dict = {'assignee_name': f'{register[7]}', 'total_actions': f'{register[20]}', 'percentage':f'{percentage}', 'zone': f'Zone {zone[0]}'}
                list_assignee_total_actions.append(assigne_dict)
            else:
                continue

    list_assignee_total_actions = sorted(list_assignee_total_actions, key = lambda item: item['total_actions'], reverse = True)
    # for item in list_total_actions_assigne:
    #     print(item)

    # Start xlsxwriter library and export all previous data generated on this file to a excel
    workbook = xlsxwriter.Workbook('Output/Inspections Data Mining AO Analytics.xlsx')

    worksheet01 = workbook.add_worksheet('Assignes, zones and actions')

    row = 1
    col = 0
    worksheet01.write(0, 0, 'assignee_name')
    worksheet01.write(0, 1, 'total_actions')
    worksheet01.write(0, 2, 'percentage')
    worksheet01.write(0, 3, 'zone')
    for assignee_total_actions in list_assignee_total_actions:
        worksheet01.write(row, col, assignee_total_actions['assignee_name'])
        worksheet01.write(row, col + 1, assignee_total_actions['total_actions'])
        worksheet01.write(row, col + 2, assignee_total_actions['percentage'])
        worksheet01.write(row, col + 3, assignee_total_actions['zone'])
        row += 1

    workbook.close()