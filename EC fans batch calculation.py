import json, requests, time
import pandas as pd
from pandas import ExcelWriter

url = "http://fanselect.net:8079/FSWebService"
user_ws, pass_ws = 'xxx', 'xxx'

# Get all the possible fans
all_results = False

def fan_ws(request_string, url):	
	ws_output = requests.post(url=url, data=request_string)
	return ws_output

def get_response(dict_request):
	dict_json = json.dumps(dict_request)
	url_response = fan_ws(dict_json, url)
	url_result = json.loads(url_response.text)
	return url_result

def sort_function(lst, n):
	lst.sort(key = lambda x: x[n])
	return lst 

# Get SessionID
session_dict = {
	'cmd': 'create_session',
	'username': user_ws,
	'password': pass_ws
}

session_id = get_response(session_dict)['SESSIONID']
print('Session ID:', session_id)
print('\n')

# Pandas import
excel_file = 'DATA_INPUT.xlsx'
df_data = pd.read_excel(excel_file)

# AHU size for merge
excel_file = 'AHU_SIZE.xlsx'
df_size = pd.read_excel(excel_file)

# Merge operation
df_data = pd.merge(df_data, df_size, on='AHU')

# IMPORTANT: INPUTS MUST BE STRINGS, NOT NUMBERS!!!!
cols = ['Height', 'Width', 'Airflow', 'Static Press.', 'No fans']
for col in cols:
	df_data[col] = df_data[col].astype(str)

df_data['RPM_static'] = 0 # ZA_N
df_data['RPM_max'] = 0 # ZA_NMAX

print('Data input:')
print(df_data.head())
print('\n')

# Check execution time
start_time = time.time()

for j in range(len(df_data['Line'])):

	line = df_data['Line'].iloc[j]
	qv = df_data['Airflow'].iloc[j]
	psf = df_data['Static Press.'].iloc[j]
	article_no = df_data['article_no'].iloc[j]
	no_fans = df_data['No fans'].iloc[j]
	height = df_data['Height'].iloc[j]
	width = df_data['Width'].iloc[j]

	# Fan request
	fan_dict = {
		'language': 'EN',
		'unit_system': 'm',
		'username': user_ws,
		'password': pass_ws,
		'cmd': 'select',
		'cmd_param': '0',
		'zawall_mode': 'ZAWALL_PLUS',
		'zawall_size': no_fans,
		'qv': qv,
		'psf': psf,
		'spec_products': 'PF_00',
		'article_no': article_no,
		'current_phase': '3',
		'voltage': '230',
		'nominal_frequency': '60',
		'installation_height_mm': height,
		'installation_width_mm': width,
		'installation_length_mm': '2000',
		'installation_mode': 'RLT_2017',
		'sessionid': session_id
	}

	print(fan_dict)
	print('\n')

	power_input = get_response(fan_dict)['ZA_PSYS']
	zawall_arr = get_response(fan_dict)['ZAWALL_ARRANGEMENT']
	no_fans = 1 if zawall_arr == 0 else int(zawall_arr[:2])
	n_actual = get_response(fan_dict)['ERP_N_ACTUAL']
	n_stat = get_response(fan_dict)['ERP_N_STAT']
	n_target = get_response(fan_dict)['ERP_N_TRAGET']
	za_n = get_response(fan_dict)['ZA_N']
	za_nmax = get_response(fan_dict)['ZA_NMAX']

	print('Number of line:', line)
	print('Fan found:', article_no)
	print('Power input W:', power_input)
	print('Eff. N_actual:', n_actual)
	print('Eff. N_stat:', n_stat)
	print('Eff. N_target:', n_target)
	print('Number of fans:', no_fans)
	print('RPM_static:', za_n)
	print('RPM_max:', za_nmax)
	print('\n')

	df_data['RPM_static'].iloc[j] = za_n
	df_data['RPM_max'].iloc[j] = za_nmax

	# Stop the loop
	print('Loop stopping!')
	print('\n')		

# Export to Excel
name = 'Results.xlsx'
writer = pd.ExcelWriter(name)
df_data.to_excel(writer, index = False)
writer.save()
