import pandas as pd
from openpyxl import load_workbook
import os.path
import os

# Define a function that works on many files

def cleanmdsrdta(file_name, sheet_name, start_row,start_col, districts, df_name):
    wb = load_workbook(file_name)
    sheet = wb[sheet_name]
    
    index = [0,1,2,3,4,5,6,7,8,9,10]
    columns = ['District','Time Period','Month','Total number of Maternal Deaths',
                'Direct - Postpartum Hemorrhage (PPH)','Direct - Antepartum hemorrhage (APH)',
               'Direct - Eclampsia','Direct - Sepsis','Direct - Ruptured uterus',
               'Direct - Obstructed labor','Other Direct MD causes',
               'Indirect - Anaemia','Indirect - Malaria','Indirect - HIV','Indirect - Pneumonia',
               'Indirect - Meningitis','Indirect - TB','Indirect - Hepatitis','Other Indirect MD causes',
               'Admin Factor - Communication problem between health facilities',
               'Admin Factor - Transport problem between health facilities',
               'Admin Factor - Lack of qualified staff',
               'Admin Factor - Lack of antibiotics',
               'Admin Factor - Lack of other essential obstetric drugs',
               'Admin Factor - Lack of essential equipment',
               'Admin Factor - Lack of laboratory facilities',
               'Admin Factor - Lack of availability of blood transfusion',
               'Admin Factor - Absence of trained stuff on duty',
               'HW Factor - Delay in deciding to refer',
               'HW Factor - Initial assessment incomplete',
               'HW Factor - Inadequate resuscitation',
               'HW Factor - Wrong diagnosis',
               'HW Factor - Wrong treatment',
               'HW Factor - No treatment',
               'HW Factor - Delay in starting treatment',
               'HW Factor - Inadequate monitoring',
               'HW Factor - Prolonged abnormal observation without action',
               'HW Factor - Lack of obstetric life saving skills',
               'Patient Factor - Delay in reporting to health facility',
               'Patient Factor - Lack of transport from home to health facility',
               'Patient Factor - Unsafe traditional/cultural practice',
               'Patient Factor - Refusal of treatment',
               'Patient Factor - Delay in decision making',
               'Number of births attended by SBA']
    df_name = pd.DataFrame(index=index,columns=columns)

    df = pd.DataFrame(sheet.values)
    df = df.T

    a = start_col
    for j in range(districts):
        for i in range(3):
            df_name.loc[i+j*3+1,'District'] = df[6][a+j*4]
            df_name.loc[i+j*3+1,'Time Period'] = df[3][1]
            df_name.loc[i+j*3+1,'Month'] = df[7][a+i+j*4]
        for i in range(3):
            df_name.loc[i+j*3+1,'Total number of Maternal Deaths'] = df[start_row][a+i+j*4]
            df_name.loc[i+j*3+1,'Direct - Postpartum Hemorrhage (PPH)'] = df[start_row+6][a+i+j*4]
            df_name.loc[i+j*3+1,'Direct - Antepartum hemorrhage (APH)'] = df[start_row+7][a+i+j*4]
            df_name.loc[i+j*3+1,'Direct - Eclampsia'] = df[start_row+8][a+i+j*4]
            df_name.loc[i+j*3+1,'Direct - Sepsis'] = df[start_row+9][a+i+j*4]
            df_name.loc[i+j*3+1,'Direct - Ruptured uterus'] = df[start_row+10][a+i+j*4]
            df_name.loc[i+j*3+1,'Direct - Obstructed labor'] = df[start_row+11][a+i+j*4]
            df_name.loc[i+j*3+1,'Other Direct MD causes'] = df[start_row+12][a+i+j*4]
            df_name.loc[i+j*3+1,'Indirect - Anaemia'] = df[start_row+15][a+i+j*4]
            df_name.loc[i+j*3+1,'Indirect - Malaria'] = df[start_row+16][a+i+j*4]
            df_name.loc[i+j*3+1,'Indirect - HIV'] = df[start_row+17][a+i+j*4]
            df_name.loc[i+j*3+1,'Indirect - Pneumonia'] = df[start_row+18][a+i+j*4]
            df_name.loc[i+j*3+1,'Indirect - Meningitis'] = df[start_row+19][a+i+j*4]
            df_name.loc[i+j*3+1,'Indirect - TB'] = df[start_row+20][a+i+j*4]
            df_name.loc[i+j*3+1,'Indirect - Hepatitis'] = df[start_row+21][a+i+j*4]
            df_name.loc[i+j*3+1,'Other Indirect MD causes'] = df[start_row+22][a+i+j*4]
            df_name.loc[i+j*3+1,'Admin Factor - Communication problem between health facilities'] = df[start_row+34][a+i+j*4] # checkpoint
            df_name.loc[i+j*3+1,'Admin Factor - Transport problem between health facilities'] = df[start_row+35][a+i+j*4]
            df_name.loc[i+j*3+1,'Admin Factor - Lack of qualified staff'] = df[start_row+36][a+i+j*4]
            df_name.loc[i+j*3+1,'Admin Factor - Lack of antibiotics'] = df[start_row+37][a+i+j*4]
            df_name.loc[i+j*3+1,'Admin Factor - Lack of other essential obstetric drugs'] = df[start_row+38][a+i+j*4]
            df_name.loc[i+j*3+1,'Admin Factor - Lack of essential equipment'] = df[start_row+39][a+i+j*4]
            df_name.loc[i+j*3+1,'Admin Factor - Lack of laboratory facilities'] = df[start_row+40][a+i+j*4]
            df_name.loc[i+j*3+1,'Admin Factor - Lack of availability of blood transfusion'] = df[start_row+41][a+i+j*4]
            df_name.loc[i+j*3+1,'Admin Factor - Absence of trained stuff on duty'] = df[start_row+42][a+i+j*4]
            df_name.loc[i+j*3+1,'HW Factor - Delay in deciding to refer'] = df[start_row+44][a+i+j*4]
            df_name.loc[i+j*3+1,'HW Factor - Initial assessment incomplete'] = df[start_row+45][a+i+j*4]
            df_name.loc[i+j*3+1,'HW Factor - Inadequate resuscitation'] = df[start_row+46][a+i+j*4]
            df_name.loc[i+j*3+1,'HW Factor - Wrong diagnosis'] = df[start_row+47][a+i+j*4]
            df_name.loc[i+j*3+1,'HW Factor - Wrong treatment'] = df[start_row+48][a+i+j*4]
            df_name.loc[i+j*3+1,'HW Factor - No treatment'] = df[start_row+49][a+i+j*4]
            df_name.loc[i+j*3+1,'HW Factor - Delay in starting treatment'] = df[start_row+50][a+i+j*4]
            df_name.loc[i+j*3+1,'HW Factor - Inadequate monitoring'] = df[start_row+51][a+i+j*4]
            df_name.loc[i+j*3+1,'HW Factor - Prolonged abnormal observation without action'] = df[start_row+52][a+i+j*4]
            df_name.loc[i+j*3+1,'HW Factor - Lack of obstetric life saving skills'] = df[start_row+53][a+i+j*4]
            df_name.loc[i+j*3+1,'Patient Factor - Delay in reporting to health facility'] = df[start_row+55][a+i+j*4]
            df_name.loc[i+j*3+1,'Patient Factor - Lack of transport from home to health facility'] = df[start_row+56][a+i+j*4]
            df_name.loc[i+j*3+1,'Patient Factor - Unsafe traditional/cultural practice'] = df[start_row+57][a+i+j*4]
            df_name.loc[i+j*3+1,'Patient Factor - Refusal of treatment'] = df[start_row+58][a+i+j*4]
            df_name.loc[i+j*3+1,'Number of births attended by SBA'] = df[start_row+67][a+i+j*4]

    df_name.loc[0,'District'] = df[6][2]
    df_name.loc[0,'Time Period'] = df[3][2]
    df_name.loc[0,'Month'] = df[7][2]
    df_name.loc[0,'Total number of Maternal Deaths'] = df[start_row][2]
    df_name.loc[0,'Direct - Postpartum Hemorrhage (PPH)'] = df[start_row+6][2]
    df_name.loc[0,'Direct - Antepartum hemorrhage (APH)'] = df[start_row+7][2]
    df_name.loc[0,'Direct - Eclampsia'] = df[start_row+8][2]
    df_name.loc[0,'Direct - Sepsis'] = df[start_row+9][2]
    df_name.loc[0,'Direct - Ruptured uterus'] = df[start_row+10][2]
    df_name.loc[0,'Direct - Obstructed labor'] = df[start_row+11][2]
    df_name.loc[0,'Other Direct MD causes'] = df[start_row+12][2]
    df_name.loc[0,'Indirect - Anaemia'] = df[start_row+15][2]
    df_name.loc[0,'Indirect - Malaria'] = df[start_row+16][2]
    df_name.loc[0,'Indirect - HIV'] = df[start_row+17][2]
    df_name.loc[0,'Indirect - Pneumonia'] = df[start_row+18][2]
    df_name.loc[0,'Indirect - Meningitis'] = df[start_row+19][2]
    df_name.loc[0,'Indirect - TB'] = df[start_row+20][2]
    df_name.loc[0,'Indirect - Hepatitis'] = df[start_row+21][2]
    df_name.loc[0,'Other Indirect MD causes'] = df[start_row+22][2]
    df_name.loc[0,'Admin Factor - Communication problem between health facilities'] = df[start_row+34][2] # checkpoint
    df_name.loc[0,'Admin Factor - Transport problem between health facilities'] = df[start_row+35][2]
    df_name.loc[0,'Admin Factor - Lack of qualified staff'] = df[start_row+36][2]
    df_name.loc[0,'Admin Factor - Lack of antibiotics'] = df[start_row+37][2]
    df_name.loc[0,'Admin Factor - Lack of other essential obstetric drugs'] = df[start_row+38][2]
    df_name.loc[0,'Admin Factor - Lack of essential equipment'] = df[start_row+39][2]
    df_name.loc[0,'Admin Factor - Lack of laboratory facilities'] = df[start_row+40][2]
    df_name.loc[0,'Admin Factor - Lack of availability of blood transfusion'] = df[start_row+41][2]
    df_name.loc[0,'Admin Factor - Absence of trained stuff on duty'] = df[start_row+42][2]
    df_name.loc[0,'HW Factor - Delay in deciding to refer'] = df[start_row+44][2]
    df_name.loc[0,'HW Factor - Initial assessment incomplete'] = df[start_row+45][2]
    df_name.loc[0,'HW Factor - Inadequate resuscitation'] = df[start_row+46][2]
    df_name.loc[0,'HW Factor - Wrong diagnosis'] = df[start_row+47][2]
    df_name.loc[0,'HW Factor - Wrong treatment'] = df[start_row+48][2]
    df_name.loc[0,'HW Factor - No treatment'] = df[start_row+49][2]
    df_name.loc[0,'HW Factor - Delay in starting treatment'] = df[start_row+50][2]
    df_name.loc[0,'HW Factor - Inadequate monitoring'] = df[start_row+51][2]
    df_name.loc[0,'HW Factor - Prolonged abnormal observation without action'] = df[start_row+52][2]
    df_name.loc[0,'HW Factor - Lack of obstetric life saving skills'] = df[start_row+53][2]
    df_name.loc[0,'Patient Factor - Delay in reporting to health facility'] = df[start_row+55][2]
    df_name.loc[0,'Patient Factor - Lack of transport from home to health facility'] = df[start_row+56][2]
    df_name.loc[0,'Patient Factor - Unsafe traditional/cultural practice'] = df[start_row+57][2]
    df_name.loc[0,'Patient Factor - Refusal of treatment'] = df[start_row+58][2]
    df_name.loc[0,'Patient Factor - Delay in decision making'] = df[start_row+59][2]
    df_name.loc[0,'Number of births attended by SBA'] = df[start_row+67][2]
    
    return df_name
