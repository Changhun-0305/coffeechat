import xlrd

loc = "/Users/changhunlee/Dropbox/CIS/CoffeeChat/coffeechat/coffeechat.xlsx"

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
	user_dict[names[i]] = list_info[i]

for (key, item) in user_dict.items():
	print(key)
	print(item[3])