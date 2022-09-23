
#'''
# This version #4 is modified based on Carina's advice
# adding feature of Dynamic % Change
# '''




# import required packages
import pandas as pd
import numpy as np
import datetime
from pywebio import start_server
from pywebio.input import *
from pywebio.output import *
import os
import warnings
warnings.filterwarnings("ignore")

# selectable years - maybe drop down list
# fiscal year / drop a normal week / weighted
# % change on In-Store Execution

# os.chdir("C:\\Users\\mdjij\\OneDrive - IRI\\Desktop\\Automation\\DF Live Session")
# os.chdir("/Users/chenjiajian/Library/CloudStorage/OneDrive-IRI/Desktop/Automation/DF Live Session")
os.chdir('C:\\Users\\mdjij\\OneDrive - IRI\\Desktop\\Automation\\Final Presentation\\Live Session')
# ignore warnings


def main():
    put_markdown("# **Forecast Adjustment Live Session**")
    put_html('<hr>')
    put_markdown("""Upload a driver forecast file downloaded from AA Suite \n 
    Please make sure you put this py. file and the driver forecast file in the same directory""")
    # Upload the driver forecast file
    file = file_upload("Upload File as required", required=True)
    # read the original_plan tab
    data = pd.read_excel(file['filename'], sheet_name= "Original_Plan")
    put_text(f"File Name: {file['filename']}")
    week0 = data.Week[0].strftime('%m-%d-%Y')
    week1 = data.Week[len(data) - 1].strftime('%m-%d-%Y')
    put_html('<br>')
    put_markdown(f"Data start from {week0} to {week1} \n")

    cols = checkbox("Structure your summary table", list(data.columns)) # what you expect to see on summary table
    cols_change = [k for k in cols if "sales" not in k.lower()] # excluding sales
    # create year column * calendar effect
    data['year'] = data.Week.apply(lambda x: x.year)
    years = list(data['year'].unique())
    # selectable years - drop down list
    fiscalyear0 = input("What fiscal year does the first row of data belong to as the first week?", type=NUMBER)
    # create a new column for "fiscal year"
    data = create_fiscalyear(int(fiscalyear0), years, data)


    year0 = input("Which year is the start year you wanna look into?", type = NUMBER, help_text = "e.g. 2020, 2021")
    year1 = input("Which year is the end year you wanna look into?", type = NUMBER, help_text = "e.g. 2023, 2024")

    # Input Assumptions
    IsDynamic = select(label="Do you want the assumptions dynamic or constant?", options= ["Constant", "Dynamic"])

    if IsDynamic == "Constant":
        Division, variables = Constant_pctChangeList(cols)
    else:
        Division, variables = Dynamic_pctChangeList(cols_change, year0, year1)
        # Division implies how does the forecasted period got divided

# Debug at 1:37 PM 08102022
    if IsDynamic != "Constant":
        data = create_period(Division, data)



    coefs = coefsList(cols_change) # dictionary - key = var name - value = coefs


    # Show the assumptions
    show_assumptionsNcoefs(IsDynamic, Division, variables, coefs, year0, year1)
    period = year1 - year0 + 1



    # TESTING
    # put_table([variables[k][y] for k in variables.keys() for y in variables[k].keys()])

    # Create colume <type>
    data['type'] = np.nan
    today = datetime.date.today()
    for k in range(len(data)):
        if data['year'][k] <= today.year:
            data['type'][k] = "Historical"
        else:
            data['type'][k] = "Forecasted"

    # create new variables for those not in the todrop list
    if IsDynamic == "Constant":
        data = create_new_variables_constant(data, cols, variables)
    else:
        data = create_new_variables_dynamic(data, cols_change, variables, Division)



    data = calculate_new_sales(data, cols_change, coefs)
    data = calculate_new_dollar(data)


    # fiscal year column
    new_cols = ['Volume Sales', 'Dollar Sales']
    for i in cols_change:
        new_cols.append(i)

    cols = new_cols

    summary_before = SummaryFrame(data, cols, year0, year1)
    summary_before = resetIndex(summary_before, year0, year1)

    cols_new = ["New " + k for k in cols]

    summary_after = SummaryFrame(data, cols_new, year0, year1)
    summary_after = resetIndex(summary_after, year0, year1)

    put_markdown("### Before The Assumptions Above")
    put_html(summary_before.to_html(border=0)).send()

    put_markdown("### After The Assumptions Above")
    put_html(summary_after.to_html(border=0)).send()

    save = actions("Save the summary table?", buttons = ["Yes","No"])
    if save == "Yes":
        df_tosave = pd.concat([summary_before,summary_after])
        df_tosave.to_excel(f"{today}.xlsx")

def create_new_variables_constant(data, cols, variables):
    for c in cols:
        c_noSpace = c.replace(" ", "_")
        data['New ' + c] = np.nan
        if "sales" not in c.lower():
            for r in range(len(data)):
                if data.loc[r,'type'] == "Historical":
                    data.loc[r,'New ' + c] = data.loc[r,c]
                else:
                    data.loc[r,'New ' + c] = data.loc[r,c] * (1+variables[c][c_noSpace])
        elif "sales" in c.lower():
            for r in range(len(data)):
                data.loc[r,'New ' + c] = data.loc[r,c]
    return data

def create_new_variables_dynamic(data, cols_change, variables, Division):
    for c in cols_change:
        data["New " + c] = np.nan
        if Division == ["Quarterly"]:
            put_markdown("to be completed...")
            # for r in range(len(data)):
                # if data.loc[r,'type'] == "Historical":
                #     data.loc[r,"New "+c] = data.loc[r,c]
                # else:
                #     for n in [i - 2022 for i in set(data[data["type"] == "Forecasted"]['fiscalyear'])]:
                #         if r <= len(data[data["type"] == "Historical"]) + (n-1)*52 + 13:
                #             data.loc[r, "New " + c] = data.loc[r, c] * variables[c][list(variables[c].keys())[0]]
                #         elif r <= len(data[data["type"]=="Historical"]) + (n-1)*52 + 26:
                #             data.loc[r, "New " + c] = data.loc[r, c] * variables[c][list(variables[c].keys())[1]]
                #         elif r <= len(data[data["type"]=="Historical"]) + (n-1)*52 + 39:
                #             data.loc[r, "New " + c] = data.loc[r, c] * variables[c][list(variables[c].keys())[2]]
                #         elif r <= len(data[data["type"]=="Historical"]) + (n-1)*52 + 52:
                #             data.loc[r, "New " + c] = data.loc[r, c] * variables[c][list(variables[c].keys())[3]]
        elif Division == ["Half"]:
            for r in range(len(data)):
                if data.loc[r,'type'] == "Historical":
                    data.loc[r,'New '+c] = data.loc[r,c]
                else:
                    data.loc[r,"New "+c] = data.loc[r,c]* variables[c][str(data.loc[r,"FiscalYear"])+"_"+ str(data.loc[r,"period"])]
    return data

# input coefs
def calculate_new_sales(d, cols_change, coefs):
    for r in range(len(d)):
        for c in cols_change:
            d.loc[r,"New Volume Sales"] =  d.loc[r,"New Volume Sales"] + coefs[c] * (d.loc[r,"New "+c] - d.loc[r,c])
            # ***
    return d


def calculate_new_dollar(d):
    for r in range(len(d)):
        d.loc[r,"New Dollar Sales"] = d.loc[r,"Price per Volume"] * d.loc[r, "New Volume Sales"]
    return d


def create_fiscalyear(x,years,data):
    l = []
    for k in range(len(years)):
        l += [x+k]*52
    l = l[:len(data)]
    data['FiscalYear'] = l
    return data

def create_period(Division, data):
    l = []
    if Division == ['Half']:
        periodperyear = ["1st"]*26 + ["2nd"]*26
        l = periodperyear * int(np.ceil(len(data)/52))
        l = l[:len(data)]
        data['Period'] = l
    if Division == ["Quarterly"]:
        periodperyear = ["1st"] * 13 + ["2nd"] * 13 + ["3rd"] * 13 + ["4th"] * 13
        l = periodperyear * int(np.ceil(len(data)/52))
        l = l[:len(data)]
        data['Period'] = l
    return data

def Constant_pctChangeList(cols):
    variables = {}
    for c in cols:
        c_woSpace = c.replace(" ", "_")
        if "sales" not in c.lower():
            variables[c] = input_group("Assumptions",
                                       [input(f"Assumptions: Input {c}'s % change in decimal", name = f"{c_woSpace}", type = FLOAT, required= True)])
    return None, variables

def Dynamic_pctChangeList(cols_change, year0, year1):
    # data = input_group("Basic info", [
    #     input('Input your name', name='name'),
    #     input('Repeat your age', name='age', type=NUMBER)
    # ], validate=check_form)
    # break it down to each year - dictionary - keys = "2021 Q1" - values = input
    thisyear = datetime.date.today().year
    Halves = ['1st', '2nd']
    Quarters = ['1st','2nd','3rd','4th']
    Years = range(year0, year1+1)
    variables = {}
    IsQuarter = checkbox("How would you like dividing a year?", options= ["Half","Quarterly"], inline = True )
    # if we divide a year by quarters
    if IsQuarter == ["Quarterly"]:
        for c in cols_change:
            variables[c] = input_group(f"{c}: ",
                    [input(f"{c}'s {y} {q} %change", name = f"{y}{q}", required = True, type = FLOAT) for y in Years if y > thisyear for q in Quarters])
    # if we divide a year by halves
    else:
        for c in cols_change:
            variables[c] = input_group(f"{c}: ",
                [input(f"{c}'s {y} {h} %change", name = f"{y}_{h}", required = True , type = FLOAT) for y in Years if y > thisyear for h in Halves])
    return IsQuarter, variables

def coefsList(cols_change):
    coefs = {}
    for c in cols_change:
        coefs[c] = input(f"Coefficients: Input {c}'s coefficient", type = FLOAT, required = True)
    return coefs


def SummaryFrame(data,cols,year0,year1):
    # first n rows: Sum
    # following with next year growth rate
    # following with CAGR
    dic = {}

    for c in cols:
        dic[c] = []
        for y in range(year0,year1+1):
            if "sales" in c.lower():
                dic[c].append(round(((data[data['FiscalYear'] == y][c].sum())/1000000),2))
            elif "price" in c.lower():
                dic[c].append(round((data[data['FiscalYear'] == y][c].mean()),2))
            else:
                dic[c].append(round((data[data['FiscalYear'] == y][c].mean()),1))
        for l in range(1,len(dic[c])):
            dic[c].append(f"{round(((dic[c][l]/dic[c][l-1] -1)*100),1)}%")
        dic[c].append(f"{round((((dic[c][year1-year0]/dic[c][0])**(1/(year1-year0))-1)*100),1)}%")
        for y in range(0,year1-year0+1):
            if "price" in c.lower():
                dic[c][y] = f"$ {dic[c][y]}"

    d = pd.DataFrame(dic)
    return d

def resetIndex(summarytable,year0,year1):
    i = list(range(year0,year1+1))
    for marker in range(year0+1,year1+1):
        i.append(f"{marker} vs {marker-1}")
    i.append(f"CAGR {year1}vs{year0}")
    summarytable.index = i
    return summarytable

def show_assumptionsNcoefs(IsDynamic, Division, variables, coefs, year0, year1):
    Years = range(year0, year1+1)
    thisyear = datetime.date.today().year
    Halves = ['1st', '2nd']
    Quarters = ['1st','2nd','3rd','4th']

    put_markdown("**The assumptions are as listed:** ")
    if IsDynamic == "Constant":
        for k in variables.keys():
            k_noSpace = k.replace(" ","_")
            if "sales" not in k.lower():
                put_markdown(f"\t**{k}**: {variables[k][k_noSpace]*100}%")
    elif IsDynamic == "Dynamic" and Division == ["Quarterly"] :
        for c in variables.keys(): # column
            put_markdown(f"*{c}:*")
            for y in Years:
                for q in range(0,4):
                    put_markdown(f"{y} {Quarters[q]}: {variables[c][y][q]}")
    elif IsDynamic == 'Dynamic' and Division == ["Half"]:
        temp = [f"{y}_{h}" for y in Years if y > thisyear for h in Halves]
        for c in variables.keys():
            put_markdown(f"*{c}:*")
            for t in temp:
                put_markdown(f"{t}: {variables[c][t]}")

    put_markdown("**The coefficients are as listed:** ")
    for coef in coefs.keys():
        put_markdown(f"\t**{coef}**: {coefs[coef]}")




if __name__ == "__main__":
    start_server(main, port=6060, debug=True)


