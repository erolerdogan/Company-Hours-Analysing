import pandas as pd

week4 = pd.read_excel("4_week.xlsx")  # Column 13 is the data which we can use
week3 = pd.read_excel("3_week.xlsx")  # Column 13 is the data which we can use
Jan_1week = pd.read_excel("Jan_1-week.xlsx")  # Column 15 is the data which we can use


def create_table(data):
    # Variables
    must_time = 2700  # workhour of a week, 540mins per day.
    dictionary = {}
    count = 0
    times = []
    numberOfDay = []

    name_list = [i for i in data[data.columns[1]]]
    for j in range(0, len(data)):
        if j < (len(data) - 1) and name_list[j] == name_list[j + 1]:
            times.append(data[data.columns[3]][j])
            count += 1

        else:
            times.append(data[data.columns[3]][j])
            count += 1
            numberOfDay.append(count)  # to count the day that the personal came to office
            count = 0
            dictionary[data[data.columns[1]][j]] = sum(times)
            times = []

    # to count the day that the personal didn't come to office
    absenceOfDay = [(5 - days) for days in numberOfDay]

    # to create new DataFrame for the days that the personals came
    df2 = pd.DataFrame(numberOfDay)

    # to create new DataFrame for the days that the personals came
    df3 = pd.DataFrame(absenceOfDay)

    new_data = pd.DataFrame(list(dictionary.items()), columns=['Personal', 'Dışarıda GeçenSüre(dk)'])
    new_data["Gün"] = df2
    new_data["EksikGün"] = df3

    # to calculate continuity rate
    q = 0
    continuity = []
    for i in new_data[new_data.columns[1]]:
        qw = (1 - ((i + (new_data[new_data.columns[3]][q]) * 540) / must_time)) * 100
        q += 1
        continuity.append(round(qw, 2))

    df4 = pd.DataFrame(continuity)
    new_data["Devamlılık(%)"] = df4
    sorted_data = new_data.sort_values(by='Devamlılık(%)')
    print(sorted_data.reset_index(drop=True))


def separate_department(department, week):
    # To create new dataframe
    department = week[week[week.columns[2]] == department]

    department = pd.concat([department[department.columns[:3]], department[department.columns[13]]], axis=1, )

    time = week[week.columns[3]]
    print("The data's date between '{}' and '{}' \n".format(time.min(), time.max()))

    # Filling the missed datas as 0
    department[department.columns[3]] = department[department.columns[3]].fillna(value=0)

    # to set index of the dataframe
    labels = [i for i in range(0, len(department))]
    department = department.set_index([pd.Series(labels)])

    create_table(department)
    print("--------------------------------------------------------------------------------------------\n\n")


def department_analysis(week):
    departments = list(set(week[week.columns[2]]))
    for i in departments:
        print("Department's name: {} \n".format(i))
        separate_department(department=i, week=week)


def company_analysis(parameter):
    # To create new dataframe
    data = pd.concat([parameter[parameter.columns[:3]], parameter[parameter.columns[13]]], axis=1)
    # Filling the missed datas as 0
    data[data.columns[3]] = data[data.columns[3]].fillna(value=0)

    x = parameter[parameter.columns[3]]
    print("The data's date between '{}' and '{}' \n".format(x.min(), x.max()))
    create_table(data)


company_analysis(week4)
department_analysis(week4)