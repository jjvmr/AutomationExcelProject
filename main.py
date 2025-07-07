import openpyxl

inv_file = openpyxl.load_workbook("coalpublic2013.xlsx")
mine_list = inv_file['Hist_Coal_Prod']

mines_per_supplier_region = {}
total_productivity_per_region = {}


# Listing each respective mine company with their respective Supplier Region
for mine_row in range(5, mine_list.max_row + 1):
    supplier_region = mine_list.cell(mine_row, 14).value
    production = mine_list.cell(mine_row, 15).value
    labour_hours = mine_list.cell(mine_row, 17).value

    # calculation of number of mining companies per region
    if supplier_region in mines_per_supplier_region:
        current_num_mines = mines_per_supplier_region.get(supplier_region)     # getting value of a key
        mines_per_supplier_region[supplier_region] = current_num_mines + 1     # setting value of a key
    else:
        mines_per_supplier_region[supplier_region] = 1


    # calculation total value of labour productivity per region (production per labour hours)
    # dividing total production value by labour hours
    if supplier_region in total_productivity_per_region:
        current_total_value = total_productivity_per_region.get(supplier_region)
        total_productivity_per_region[supplier_region] = current_total_value + (production / labour_hours)
    elif production == 0:
        total_productivity_per_region[supplier_region] = production + labour_hours
    else:
        total_productivity_per_region[supplier_region] = production / labour_hours