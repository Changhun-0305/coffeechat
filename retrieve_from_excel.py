import xlrd

loc = "CoffeeChatSpring.xlsx"

if __name__ == "__main__":
	wkbk = xlrd.open_workbook(loc)
	sheet = wkbk.sheet_by_index(0)

	num_user = sheet.nrows

	names = sheet.col_values(0)
	grades = sheet.col_values(1)
	fun_facts = sheet.col_values(2)
	list_set_ban = []
	for i in range(num_user):
		s = set(sheet.cell_value(i, 3).split(", "))
		list_set_ban.append(s)

	numbers = sheet.col_values(4)

	list_info = []
	for i in range(num_user):
		info_per_row = []
		info_per_row.append(grades[i])
		info_per_row.append(fun_facts[i])
		info_per_row.append(list_set_ban[i])
		info_per_row.append('+1' + str(int(numbers[i])))
		list_info.append(info_per_row)

	user_dict = dict()
	for i in range(num_user):
		info = list_info[i]
		indices = [names.index(x) for x in info[2] if x in names]
		info.append(indices)
		user_dict[names[i]] = list_info[i]