from Modules.FSI.fsi_inspections_data_mining_export_excel import inspection_data_mining_fsi
from Modules.FSI.fsi_question_level_export_excel import question_level_fsi
from Modules.AO.ao_inspections_data_mining_export_excel import inspection_data_mining_ao
from Modules.AO.ao_question_level_export_excel import question_level_ao

inspection_data_mining_ao_csv = 'Input/Arquivos AO/Inspections Data Mining AO.csv'
zones_csv = 'Input/Arquivos AO/Zones.csv'
question_level_ao_csv = 'Input/Arquivos AO/Question level AO.csv'
questions_categories_ao_csv = 'Input/Arquivos AO/Questions categories AO.csv'
inspection_data_mining_fsi_csv = 'Input/Arquivos FSI/Inspections Data Mining FSI.csv'
question_level_fsi_csv = 'Input/Arquivos FSI/Question level FSI.csv'
questions_categories_fsi_csv = 'Input/Arquivos FSI/Questions categories FSI.csv'

try:
	inspection_data_mining_fsi(inspection_data_mining_fsi_csv)
except FileNotFoundError:
    print(f"Couldn't find the {inspection_data_mining_fsi_csv}")
else:
    pass

try:
	question_level_fsi(question_level_fsi_csv, questions_categories_fsi_csv)
except FileNotFoundError:
    print(f"Couldn't find the {question_level_fsi_csv}")
else:
    pass

try:
	inspection_data_mining_ao(inspection_data_mining_ao_csv, zones_csv)
except FileNotFoundError:
    print(f"Couldn't find the {inspection_data_mining_ao_csv}")
else:
    pass

try:
	question_level_ao(question_level_ao_csv, questions_categories_ao_csv)
except FileNotFoundError:
    print(f"Couldn't find the {question_level_ao_csv}")
else:
    pass