import os
import openpyxl
import pandas as pd
import numpy as np
import date_processing

employee_file = 'data/all_members.xlsx'
allocation_file = 'data/allocation_report.xlsx'
def date_processing():
    all_date = pd.DataFrame({'all_date': pd.date_range('2023-01-01', '2023-12-31')})

    # bỏ các ngày cuối tuần
    weekend_cond_ = (all_date['all_date'].dt.dayofweek != 5) & (all_date['all_date'].dt.dayofweek != 6)
    all_date = all_date[weekend_cond_] 

    all_date['Month'] = all_date['all_date'].dt.month
    all_date['Week'] = all_date['all_date'].dt.isocalendar().week

    ref_date = pd.DataFrame()

    for month in range(1, 13):
        cond_ = all_date['Month'] == month
        each_month = all_date[cond_].copy()
        each_month['Week number'] = each_month['Week'] - (each_month['Week'].min() - 1)
        ref_date = pd.concat([ref_date, each_month])

    return ref_date

def employee_processing(employee_file, ref_date):
    all_emp = pd.read_excel(employee_file)
    #mặc định ngày tháng
    all_emp['In SDU From'] = all_emp['In SDU From'].fillna(pd.to_datetime('2023-01-01'))
    all_emp['In SDU To'] = all_emp['In SDU To'].fillna(pd.to_datetime('2023-12-31'))

    all_emp['in_sdu_time'] = [list(pd.date_range(start, end)) for start, end in zip(all_emp['In SDU From'], all_emp['In SDU To'])]
    all_emp_all_date = all_emp[['Acc', 'in_sdu_time']].explode('in_sdu_time')

    off_emp = all_emp[all_emp['Off Type'].notnull()].copy().reset_index()
    off_emp['off_duration'] = [list(pd.date_range(start, end)) for start, end in zip(off_emp['Off From'], off_emp['Off To'])]
    all_emp_off_date = off_emp.loc[:, ['Acc', 'off_duration']].copy()
    all_emp_off_date = all_emp_off_date.explode('off_duration')
    all_emp_work_days = pd.merge(all_emp_all_date, all_emp_off_date
                                 , how='left'
                                 , left_on=['Acc', 'in_sdu_time']
                                 , right_on=['Acc', 'off_duration'])

    all_emp_work_days = all_emp_work_days.loc[all_emp_work_days['off_duration'].isnull(), ['Acc', 'in_sdu_time']].copy()

    all_emp_work_days.rename(columns={'in_sdu_time': 'Working Day'}, inplace=True)
    all_emp_work_days.reset_index(drop=True, inplace=True)
    all_emp_work_days = pd.merge(all_emp_work_days, ref_date
                                 , how='inner'
                                 , left_on=['Working Day']
                                 , right_on=['all_date']
                                 )


    return all_emp_work_days


def alloctation_processing(alloctation_file):
        allc = pd.read_excel(alloctation_file)
        allc = allc.loc[allc['Username'].isnull() == False, :].copy().reset_index(drop=True)
        allc['project_duration'] = [[_ for _ in pd.date_range(allc['From Date'][i], allc['To Date'][i])] for i in
                                    range(allc.shape[0])]
        allc = allc.explode('project_duration')
        return allc

def employee_information(employee_file):
        all_emp = pd.read_excel(employee_file)
        emp_inf= all_emp.loc[:,['Name','Acc','Branch','Job Title']].copy()
        return  emp_inf

def allocation_by_week(allc, all_emp_work_days):
        in_pj_days = allc.groupby(['Username', 'project_duration'])['Hours / Day'].sum().reset_index()
        in_pj_days['Hours / Day'].unique()
        allc_date = pd.merge(all_emp_work_days, in_pj_days
                             , how='left'
                             , left_on=['Acc', 'Working Day']
                             , right_on=['Username', 'project_duration'])

        allc_date['Hours / Day'] = allc_date['Hours / Day'].fillna(0)

        allc_date.drop(columns=['Username', 'project_duration', 'all_date'], inplace=True)

        hours_by_week = allc_date.groupby(['Acc', 'Month', 'Week number'])['Hours / Day'].sum().reset_index()
        days_by_week = allc_date.groupby(['Acc', 'Month', 'Week number'])['Working Day',].count().reset_index()

        allc_by_week = pd.merge(hours_by_week, days_by_week
                                , how='inner'
                                , on=['Acc', 'Month', 'Week number']
                                )
        #Hiệu suất làm việc trong tuần
        allc_by_week['Calendar Effort'] = allc_by_week['Hours / Day'] / (allc_by_week['Working Day'] * 8)

        #gán giá trị theo từng hiệu suất
        cond_list = [allc_by_week['Calendar Effort'] == 0
                    ,allc_by_week['Calendar Effort'] == 1
                    ]
        values = ['IDLE', 'Allocated']

        allc_by_week['Calendar Effort Type'] = np.select(cond_list, values, 'Not full Allocated')

        return allc_by_week

def calculate_allocated_number_by_week(allc_by_week):
    allocated = allc_by_week[allc_by_week['Calendar Effort Type'] == 'Allocated'].copy()
    allocated_by_week = allocated.groupby(['Month', 'Week number'])['Acc'].nunique().reset_index()
    return allocated_by_week


def caculate_not_full_allocated_number_by_week(allc_by_week):
        not_f_allocated = allc_by_week.loc[allc_by_week['Calendar Effort Type'] == 'Not full Allocated', :].copy()
        not_f_allocated_by_week = not_f_allocated.groupby(['Month', 'Week number'])['Acc'].nunique().reset_index()
        return not_f_allocated_by_week

def caculate_idle_number_by_week(allc_by_week):
        idle = allc_by_week.loc[allc_by_week['Calendar Effort Type'] == 'IDLE', :].copy()
        idle_by_week = idle.groupby(['Month', 'Week number'])['Acc'].nunique().reset_index()
        return  idle_by_week



def information_by_user(allc, ref_date, all_emp, allc_by_week, month, week_number):
    #thời gian thực hiện dự án
    pj_by_week = pd.merge(allc, ref_date, how='inner', left_on=['project_duration'], right_on=['all_date'])

    MONTH = month
    WEEK_NUMBER = week_number

    #từng dự án được thực hiện trong tháng và tuần theo yc
    get_list_row_cond = (pj_by_week['Month'] == MONTH) & (pj_by_week['Week number'] == WEEK_NUMBER)
    get_list_by_week = pj_by_week[get_list_row_cond].copy()

    #tạo df theo Calendar Effort Type
    allocated = allc_by_week[allc_by_week['Calendar Effort Type'] == 'Allocated'].copy()
    not_f_allocated = allc_by_week[allc_by_week['Calendar Effort Type'] == 'Not full Allocated'].copy()
    idle = allc_by_week[allc_by_week['Calendar Effort Type'] == 'IDLE'].copy()

    #list Type
    allocated_acc_cond = allocated.loc[(allocated['Month'] == MONTH) & (allocated['Week number'] == WEEK_NUMBER), 'Acc'].unique()
    not_full_allocated_acc_cond = not_f_allocated.loc[(not_f_allocated['Month'] == MONTH) & (not_f_allocated['Week number'] == WEEK_NUMBER), 'Acc'].unique()
    idle_acc_cond = idle.loc[(idle['Month'] == MONTH) & (idle['Week number'] == WEEK_NUMBER), 'Acc'].unique()

    # Align DataFrames

    allocated_list = get_list_by_week[get_list_by_week['Username'].isin(allocated_acc_cond)].copy()
    allocated_list = allocated_list.merge(all_emp, left_on='Username', right_on='Acc', how='left')

    not_f_allocated_list = get_list_by_week[get_list_by_week['Username'].isin(not_full_allocated_acc_cond)].copy()
    idle_list = all_emp[all_emp['Acc'].isin(idle_acc_cond)].copy()

    return allocated_list, not_f_allocated_list, idle_list


def get_total_employee_by_week(all_emp_work_days, MONTH, WEEK_NUMBER):
            #tổng số nv làm việc trong tuần thnasg theo yc
            total_emp_by_week = all_emp_work_days.groupby(['Month','Week number'])['Acc'].nunique().reset_index()
            return total_emp_by_week.loc[(total_emp_by_week['Month']==MONTH) & (total_emp_by_week['Week number']==WEEK_NUMBER),:].copy()

def get_allocated_number_by_specify_week(get_allocated_number_by_week, allc_by_week,MONTH,WEEK_NUMBER):
            #nv đc phân bổ công việc theo tháng và tuần
            get_allocated_number_by_week = get_allocated_number_by_week(allc_by_week)
            return get_allocated_number_by_week.loc[(get_allocated_number_by_week['Month']==MONTH) & (get_allocated_number_by_week['Week number']==WEEK_NUMBER),:].copy()

def get_not_f_allocated_number_by_specify_week(caculate_not_full_allocated_number_by_week,allc_by_week, MONTH, WEEK_NUMBER):
            #nv chưa đc phân bổ cv
            caculate_not_full_allocated_number_by_week = caculate_not_full_allocated_number_by_week(allc_by_week)
            return  caculate_not_full_allocated_number_by_week.loc[(caculate_not_full_allocated_number_by_week['Month']==MONTH) & (caculate_not_full_allocated_number_by_week['Week number']==WEEK_NUMBER),:].copy()

def get_idle_number_by_specify_week(caculate_idle_number_by_week, allc_by_week, MONTH, WEEK_NUMBER):
            #nv k có cv
            caculate_idle_number_by_week = caculate_idle_number_by_week(allc_by_week)
            return caculate_idle_number_by_week.loc[(caculate_idle_number_by_week['Month']==MONTH) & (caculate_idle_number_by_week['Week number']==WEEK_NUMBER),:].copy()

def get_summary_table(get_total_employee_by_week,
                      get_allocated_number_by_specify_week,
                      get_not_f_allocated_number_by_specify_week,
                      get_idle_number_by_specify_week,
                      all_emp_work_days,
                      caculate_allocated_number_by_week,
                      caculate_not_full_allocated_number_by_week,
                      allc_by_week,
                      MONTH,
                      WEEK_NUMBER,
                      caculate_idle_number_by_week):


    get_total_employee_by_week = get_total_employee_by_week

    get_allocated_number_by_specify_week = caculate_allocated_number_by_week

    get_not_f_allocated_number_by_specify_week = caculate_not_full_allocated_number_by_week

    get_idle_number_by_specify_week = caculate_idle_number_by_week

    get_total_employee_by_week['Type'] = 'Total Employees'
    get_allocated_number_by_specify_week['Type'] = 'Allocated'
    get_not_f_allocated_number_by_specify_week['Type'] = 'Not full Allocated'
    get_idle_number_by_specify_week['Type'] = 'IDLE'
    return pd.concat([get_total_employee_by_week, get_allocated_number_by_specify_week, get_not_f_allocated_number_by_specify_week, get_idle_number_by_specify_week])



