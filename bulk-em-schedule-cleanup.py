import openpyxl

# Open a specified Excel file.
filename = input("Enter the name of the Excel file: ")
workbook = openpyxl.load_workbook(filename)

# Prompt for columns to be used.
print('Loaded workbook.')
schedule_column_letter = input("Enter the column letter for the 'schedule_name' column (e.g. A, B, C, V): ")
version_column_letter = input("Enter the column letter for the 'version' column (e.g. A, B, C, AS): ")
worksheet = workbook.active
print('Workbook activated.')

# Start at the third row and define variables. Row one is a header. This avoids adding a blank row between the header and the first row.
current_row = 3
schedule_cell = worksheet[schedule_column_letter + str(current_row)]
version_cell = worksheet[version_column_letter + str(current_row)]
count = 0   # Will be used to track number of lines completed.
rows_to_delete = [] # Empty container to hold rows marked for deletion. Deletion is the slowest step, so batching it is efficient.

# Redefines the current cells after any given row change and updates the user with progress.
def new_current_cell_after_insert_delete():
    global count, schedule_cell, version_cell
    schedule_cell = worksheet[schedule_column_letter + str(current_row)]
    version_cell = worksheet[version_column_letter + str(current_row)]
    count += 1
    print(f'Evaluated {count} of 19,153 lines.')
    return count, schedule_cell, version_cell

# Iterate through the rows. Insert blank rows between schedules, delete non-current versions of each schedule. Runs until it hits an empty cell.
# PRE-CONDITION: Sheet needs to be sorted in this order: By 'schedule_name' (optionally ASC), then by 'version' DESC, then by 'unit' ASC, then optionally by 'sampling_point' ASC.
while schedule_cell.value != None:
    
    if schedule_cell.value == worksheet[schedule_column_letter + str(current_row-1)].value:
        
        highest_version_for_this_schedule = worksheet[version_column_letter + str(current_row-1)].value
        
        while schedule_cell.value == worksheet[schedule_column_letter + str(current_row-1)].value:     
            
            if version_cell.value < highest_version_for_this_schedule:
               
                rows_to_delete.append(current_row)
                print(f'Rows awaiting deletion: {rows_to_delete}')
                current_row += 1
                
            else:
                current_row += 1
             
            new_current_cell_after_insert_delete()
        
        if rows_to_delete != []:
    
            worksheet.delete_rows(idx = rows_to_delete[0], amount = len(rows_to_delete))
            current_row -= len(rows_to_delete)
            rows_to_delete = []
         
    else:
        worksheet.insert_rows(current_row)
        current_row += 2   
    
    new_current_cell_after_insert_delete()
    

# Save the modified workbook
new_filename = f'{filename[:-5]} PROPERLY FORMATTED DATA FILE.xlsx'
workbook.save(new_filename)
print('All done UwU ;3.')