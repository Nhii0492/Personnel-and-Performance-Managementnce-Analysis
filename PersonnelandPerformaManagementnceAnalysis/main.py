import os
import openpyxl
import pandas as pd
import numpy as np
import processing

def get_month():
    while True:
        try:
            month = int(input("Nhập tháng: "))
            if 1 <= month <= 12:
                return month
            else:
                print("Vui lòng nhập lại. Tháng phải nằm trong khoảng từ 1 đến 6.")
        except ValueError:
            print("Tháng là giá trị nguyên.")

def get_week_number():
    while True:
        try:
            week = int(input("Nhập tuần (1-6): "))
            if 1 <= week <= 6:
                return week
            else:
                print("Vui lòng nhập lại. Tuần phải nằm trong khoảng từ 1 đến 6.")
        except ValueError:
            print("Tuần là giá trị nguyên.")
def main():
    employee_file = 'data/all_members.xlsx'
    allocation_file = 'data/allocation_report.xlsx'
    output_file = 'data/master_report.xlsx'

    MONTH = get_month()
    WEEK_NUMBER = get_week_number()

    ref_date = processing.date_processing()
    all_emp_work_days = processing.employee_processing(employee_file, ref_date)
    emp_inf = processing.employee_information(employee_file)

    allc = processing.alloctation_processing(allocation_file)
    allc_by_week = processing.allocation_by_week(allc, all_emp_work_days)

    allocated_by_week = processing.calculate_allocated_number_by_week(allc_by_week)
    not_f_allocated_by_week = processing.caculate_not_full_allocated_number_by_week(allc_by_week)
    idle_by_week = processing.caculate_idle_number_by_week(allc_by_week)

    allocated_list_by_week, not_f_allocated_list_by_week, idle_list_by_week = processing.information_by_user(
                                                                                                    allc=allc,
                                                                                                    ref_date=ref_date,
                                                                                                    all_emp=emp_inf,
                                                                                                    allc_by_week=allc_by_week,
                                                                                                    month=MONTH,
                                                                                                    week_number=WEEK_NUMBER
                                                                                                )

    # Remove duplicate
    allocated_list_by_week.drop_duplicates(subset='Username', inplace=True)
    not_f_allocated_list_by_week.drop_duplicates(subset='Username', inplace=True)
    idle_list_by_week.drop_duplicates(subset='Acc', inplace=True)

    # gọi func
    total_employee_by_week = processing.get_total_employee_by_week(all_emp_work_days, MONTH, WEEK_NUMBER)
    allocated_number_by_week = processing.get_allocated_number_by_specify_week(processing.calculate_allocated_number_by_week, allc_by_week, MONTH, WEEK_NUMBER)
    not_f_allocated_number_by_week = processing.get_not_f_allocated_number_by_specify_week(processing.caculate_not_full_allocated_number_by_week, allc_by_week, MONTH, WEEK_NUMBER)
    idle_number_by_week = processing.get_idle_number_by_specify_week(processing.caculate_idle_number_by_week, allc_by_week, MONTH, WEEK_NUMBER)

    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # Create summary table
        summary_table = processing.get_summary_table(total_employee_by_week, allocated_number_by_week,
                                          not_f_allocated_number_by_week, idle_number_by_week, all_emp_work_days,
                                          allocated_by_week, not_f_allocated_by_week, allc_by_week, MONTH, WEEK_NUMBER,
                                          idle_by_week)

        # Filter and save the summary table
        summary_table_filtered = summary_table[
            (summary_table['Month'] == MONTH) & (summary_table['Week number'] == WEEK_NUMBER)]
        summary_table_filtered.to_excel(writer, sheet_name='Summary', index=False)

        # Save other sheets
        allocated_list_by_week.to_excel(writer, sheet_name='Allocated')
        not_f_allocated_list_by_week.to_excel(writer, sheet_name='Not full Allocated')
        idle_list_by_week.to_excel(writer, sheet_name='IDLE')


if __name__ == '__main__':
    main()
    print('Saved file successfully')



