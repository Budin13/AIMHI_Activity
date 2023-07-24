import openpyxl as xl
import pandas as pands
import json


sourceFile = pands.read_csv("source.csv")
#creating excel files
def create_file():
    fileName = ["Type.xlsx", "FormGroup.xlsx", "Report.xlsx"]
    workbook = xl.Workbook()
    #Naming xlsx files
    for naming in fileName:
        workbook.save(naming)

#getting the data of Type column w/out dulplicaes
def import_type():
    df = sourceFile["Type"].drop_duplicates().reset_index(drop=True).to_frame(name="type") 
    df.insert(0,'id', range(1, 1 + len(df))) #auto incremented values for 'id' column
    df.to_excel('Type.xlsx', index=False)

#getting the data of Report Forms Group column w/out dulplicaes
def import_form_group():
    df = sourceFile["Report Forms Group"].drop_duplicates().reset_index(drop=True).to_frame(name="form_group")
    df.insert(0,'id', range(1, 1 + len(df))) #auto incremented values for 'id' column
    df.to_excel('FormGroup.xlsx' , index=False)

def import_report():

    sourceFile.to_excel('Report.xlsx', index = None )

    # Load the 'Type.xlsx' and 'FormGroup.xlsx' create a mapping dictionary
    dictType = pands.read_excel("Type.xlsx")
    dictFormGroup = pands.read_excel("FormGroup.xlsx")
    type_mapping = dict(zip(dictType['type'], dictType['id']))
    form_group_mapping = dict(zip(dictFormGroup['form_group'], dictType['id']))

    # Read 'Report.xlsx' into a DataFrame
    dictReport = pands.read_excel("Report.xlsx")

    # Use the 'map' function to replace 'Type' and 'Report Forms Group' values with corresponding 'id' values
    dictReport['Type'] = dictReport['Type'].map(type_mapping)
    dictReport['Report Forms Group'] = dictReport['Report Forms Group'].map(form_group_mapping)

    # Save the updated DataFrame back to 'Report.xlsx'
    dictReport.to_excel("Report.xlsx", index=False)

def display_report():

    
    form_group = pands.read_excel('FormGroup.xlsx')
    dictFormGroup = form_group.set_index('id')['form_group'].to_dict()
    type = pands.read_excel('Type.xlsx')
    dicttype = type.set_index('id')['type'].to_dict()

    # Convert the data to JSON format with original values for 'Type' and 'Report Forms Group' columns
    df_report = pands.read_excel('Report.xlsx')
    report_data = df_report.to_dict(orient='records')
    json_data = []
    for row in report_data:
        json_row = {}
        for key, value in row.items():
            if key in ['Type']:
                # Check if the value exists in the form_group_dict before accessing it
                json_row[key] = dicttype.get(value, value)
            elif key in ['Report Forms Group']:
                json_row[key] = dictFormGroup.get(value, value)
            else:
                json_row[key] = value
        json_data.append(json_row)

    for report in json_data:
        print(json.dumps(report, indent=1))



#Calling the functions
create_file()
import_type()
import_form_group()
import_report()
display_report()


