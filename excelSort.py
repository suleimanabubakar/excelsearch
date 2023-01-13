from ifExists import main_filter

present_workbook = input(f'ENTER BASE FILE <: ')
present_column = input(f'ENTER COLUMN THAT CONTAINS ID NO FOR BASE <: ')
target_workbook = input(f'ENTER TARGET FILE <: ')
target_column = input(f'ENTER COLUMN THAT CONTAINS ID NO FOR TARGET <: ')


main_filter(present_workbook,target_workbook,present_column,target_column)


