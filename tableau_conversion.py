# This program organizes Dreamer FW data into a Tableau readable format.

import csv
import numpy
from enum import Enum
from sets import Set
from xlwt import Workbook, Formula
import xlwt

wb = Workbook()
style = xlwt.XFStyle()
font = xlwt.Font()
font.bold = True
style.font = font

num_dp = 10
col_after = 6
preset_data = 6

class Columns(Enum):
	last_name = 0
	first_name = 1
	ID = 2
	GR = 3
	core = 4
	w_date = 5
	w_gain = 6
	w_overall_ss = 7
	w_overal_level = 8
	w_overall_percentile = 9
	w_overall_tier = 10
	w_num_ss = 11
	w_num_level = 12
	w_num_tier = 13
	w_alg_ss = 14
	w_alg_level = 15
 	w_alg_tier = 16
	w_data_ss = 17
	w_data_level = 18
	w_data_tier = 19
	w_geo_ss = 20
	w_geo_level = 21
	w_geo_tier = 22
	f_date = 23
	f_gain = 24
	f_overall_ss = 25
	f_overal_level = 26
	f_overall_percentile = 27
	f_overall_tier = 28
	f_num_ss = 29
	f_num_level = 30
	f_num_tier = 31
	f_alg_ss = 32
	f_alg_level = 33
 	f_alg_tier = 34
	f_data_ss = 35
	f_data_level = 36
	f_data_tier = 37
	f_geo_ss = 38
	f_geo_level = 39
	f_geo_tier = 40


def ask():
	while True:
		try:
			dreamer_data = raw_input("Enter a file name: ").lower()   
			open(dreamer_data) 
		except IOError:
			print("Invalid file name; try again.")
			continue
		else:
			break
	return dreamer_data

def initialize_info(row, row_num):
	for col in range(preset_data):
		if col == 5 and ((row_num - 1) % 10) > 4: #first 3 dates in 6 pt data sets are winter, this accounts for fall.
			rocky_math.write(row_num, col, row[Columns.f_date.value])
		else:
			rocky_math.write(row_num, col, row[col])

def write_data(start_col, row_num):
	if start_col == Columns.w_overall_ss.value or start_col == Columns.f_overall_ss.value:
		for i in range (col_after, col_after + 4):
			rocky_math.write(row_num, i, row[i + 1])
	else: 
		for i in range (col_after, col_after + 3):
			rocky_math.write(row_num, i, row[i + 1])


dreamer_data = ask()
rocky_math = wb.add_sheet("Rocky Reading")
row_num = 1
with open(dreamer_data) as csvfile:
	readCSV = csv.reader(csvfile, delimiter = ',')
	for row in readCSV:
		for i in range(num_dp):
			initialize_info(row, row_num)
			if i == 0:
				write_data(Columns.w_overall_ss.value, row_num)
			elif i == 1:
				write_data(Columns.w_num_ss.value, row_num)
			elif i == 2: 
				write_data(Columns.w_alg_ss.value, row_num)
			elif i == 3:
				write_data(Columns.w_data_ss.value, row_num)
			elif i == 4:
				write_data(Columns.w_geo_ss.value, row_num)
			elif i == 5:
				write_data(Columns.f_overall_ss.value, row_num)
			elif i == 6:
				write_data(Columns.f_num_ss.value, row_num)
			elif i == 7:
				write_data(Columns.f_alg_ss.value, row_num)
			elif i == 8:
				write_data(Columns.f_data_ss.value, row_num)
			elif i == 9:
				write_data(Columns.f_geo_ss.value, row_num)
			row_num += 1

wb.save('Rocky_Data.xls')

