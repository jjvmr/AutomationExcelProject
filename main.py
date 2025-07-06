import openpyxl

inv_file = openpyxl.load_workbook("coalpublic2013.xlsx")
mine_list = inv_file['Hist_Coal_Prod']

mines_per_supplier_region = {}

# total rows: 1454
for mine_row in range(5, mine_list.max_row + 1):
    supplier_region = mine_list.cell(mine_row, 14).value

    if supplier_region in mines_per_supplier_region:
        current_num_mines = mines_per_supplier_region.get(supplier_region)     # getting value of a key
        mines_per_supplier_region[supplier_region] = current_num_mines + 1     # setting value of a key
    else:
        mines_per_supplier_region[supplier_region] = 1

print(mines_per_supplier_region)
