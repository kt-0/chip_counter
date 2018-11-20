#!/usr/local/bin/python3.7
import pandas as pd

# NOTE(s): #1 - for this to work as an executable, all file paths need to be absolute


def main():


	#1
	df = pd.read_excel('assets/excel/chip_counter.xlsx', header=0)
	df_filled = df.fillna(method='bfill', axis=0)
	df.index = pd.MultiIndex.from_arrays([df_filled['Date'], df_filled['Time']])
	df.drop(columns=['Date', 'Time'], inplace=True)

	snack_choices = {
		"1": {"name": "Doritos (Nacho Cheese)", "category": "chips"},
		"2": {"name": "Lays (Pickle)", "category": "chips"},
		"3": {"name": "Cheetos (Flamin' Hot)", "category": "chips"},
		"4": {"name": "Doritos (Cool Ranch)", "category": "chips"},
		"5": {"name": "Takis Pop (Fuego)", "category": "chips"},
		"6": {"name": "Ruffles (Cheddar & Sour Cream)", "category": "chips"},
		"7": {"name": "Wavy Lays (Hickory BBQ)", "category": "chips"},
		"8": {"name": "Pocky (Chocolate)", "category": "cookies"}
		}


	snack_vendor = []
	for x in snack_choices.items():
		name = x[1]['name']
		vend_num = x[0]
		snack_vendor.append(" [{0}]{2:^3}{1:<25}".format(vend_num, name, "-"))
		if int(vend_num)%2==0: # two choices per line
			snack_vendor[-1] = snack_vendor[-1]+"\n"

	print("".join(snack_vendor))
	selection = input("Make selection: ")
	while ((not selection.isdigit()) or (selection not in snack_choices.keys())):
		print("*************************")
		print("********* WRONG *********")
		print("*************************\n")
		print("".join(snack_vendor))
		selection = input("Selection should be a number between 1 and {} ".format(list(snack_choices.keys())[-1]))

	choice = snack_choices[selection]['name']
	print(choice)
	now = pd.MultiIndex.from_arrays([[pd.datetime.today().date()], [pd.datetime.today().time()]], names=['Date','Time'])
	dfnew = pd.DataFrame({'Snack': choice}, index=now)
	df = df.append(dfnew, sort=False)



	write_xlsx(df)


def write_xlsx(df):
	#1
	writer = pd.ExcelWriter("assets/excel/chip_counter.xlsx", engine="xlsxwriter", date_format='mmm-dd')
	df.to_excel(writer, sheet_name="Sheet1")
	workbook = writer.book
	worksheet = writer.sheets["Sheet1"]
	light_blue, dark_blue, mid_blue = '#D9E1F1', '#4674C1', '#8FAAD9'
	grey, greyish_blue = '#C9C9C9', '#b6d2ed'
	yellow, maroon = '#FED330', '#A9180F'
	base = {
		'font_name': 'AppleGothic',
		'align': 'center',
		'font_size': 10
	}
	head = {**base,
		'font_size': 14,
		'font_color': 'white',
		'bg_color': dark_blue,
		'bold': True,
		'bottom': 1,
		}
	date = {**base, 'num_format': 'mmm-dd'}
	time = {**base, 'num_format': 'hh:mm AM/PM'}

	head_format = workbook.add_format(head)
	date_format = workbook.add_format(date)
	time_format = workbook.add_format(time)

	#print(df.index.names)

# format the head
	# index columns
	for i,x in enumerate(df.index.names):
		worksheet.write(0, i, x, head_format)

	# non-index column
	for col_num, value in enumerate(df.columns.values):
		# print(df.columns)
		worksheet.write(0, 2, value, head_format)


	# writing time and date with desired format
	for i in range(1, df.shape[0]+1):

		x = i-1
		d = df.index[x][0].date()
		t = df.index[x][1]

		worksheet.write_datetime(i, 0, d, date_format)
		worksheet.write_datetime(i, 1, t, time_format)

	worksheet.set_column(0, 0, 20)
	worksheet.set_column(1, 1, 20)
	worksheet.set_column(2, 2, 20)

	writer.save()

main()
