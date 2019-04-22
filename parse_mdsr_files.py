# Define a function that works on many files

def cleanmdsrdta(file_name, sheet_name, start_col, districts, df_name):
    wb = load_workbook(file_name)
    sheet = wb[sheet_name]
    
    index = [0,1,2,3,4,5,6,7,8,9,10]
    columns = ['Zone','District','Time Period','Month','Total number of Maternal Deaths',
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
               'Patient Factor - Delay in decision making']
    df_name = pd.DataFrame(index=index,columns=columns)

    df = pd.DataFrame(sheet.values)
    df = df.T

    a = start_col
    for j in range(districts):
        for i in range(3):
            df_name.loc[i+j*3,'District'] = df[6][a+j*4]
            df_name.loc[i+j*3,'Time Period'] = df[3][1]
            df_name.loc[i+j*3,'Month'] = df[7][a+i+j*4]
        for i in range(3):
            df_name.loc[i+j*3,'Total number of Maternal Deaths'] = df[start_row][a+i+j*4]
            df_name.loc[i+j*3,'Direct - Postpartum Hemorrhage (PPH)'] = df[start_row+7][a+i+j*4]
            df_name.loc[i+j*3,'Direct - Antepartum hemorrhage (APH)'] = df[start_row+8][a+i+j*4]
            df_name.loc[i+j*3,'Direct - Eclampsia'] = df[start_row+9][a+i+j*4]
            df_name.loc[i+j*3,'Direct - Sepsis'] = df[start_row+10][a+i+j*4]
            df_name.loc[i+j*3,'Direct - Ruptured uterus'] = df[start_row+11][a+i+j*4]
            df_name.loc[i+j*3,'Direct - Obstructed labor'] = df[start_row+12][a+i+j*4]
            df_name.loc[i+j*3,'Other Direct MD causes'] = df[start_row+13][a+i+j*4]
            df_name.loc[i+j*3,'Indirect - Anaemia'] = df[start_row+15][a+i+j*4]
            df_name.loc[i+j*3,'Indirect - Malaria'] = df[start_row+16][a+i+j*4]
            df_name.loc[i+j*3,'Indirect - HIV'] = df[start_row+17][a+i+j*4]
            df_name.loc[i+j*3,'Indirect - Pneumonia'] = df[start_row+18][a+i+j*4]
            df_name.loc[i+j*3,'Indirect - Meningitis'] = df[start_row+19][a+i+j*4]
            df_name.loc[i+j*3,'Indirect - TB'] = df[start_row+20][a+i+j*4]
            df_name.loc[i+j*3,'Indirect - Hepatitis'] = df[start_row+21][a+i+j*4]
            df_name.loc[i+j*3,'Other Indirect MD causes'] = df[start_row+22][a+i+j*4]
            df_name.loc[i+j*3,'Admin Factor - Communication problem between health facilities'] = df[43][a+i+j*4]
            df_name.loc[i+j*3,'Admin Factor - Transport problem between health facilities'] = df[44][a+i+j*4]
            df_name.loc[i+j*3,'Admin Factor - Lack of qualified staff'] = df[45][a+i+j*4]
            df_name.loc[i+j*3,'Admin Factor - Lack of antibiotics'] = df[46][a+i+j*4]
            df_name.loc[i+j*3,'Admin Factor - Lack of other essential obstetric drugs'] = df[47][a+i+j*4]
            df_name.loc[i+j*3,'Admin Factor - Lack of essential equipment'] = df[48][a+i+j*4]
            df_name.loc[i+j*3,'Admin Factor - Lack of laboratory facilities'] = df[49][a+i+j*4]
            df_name.loc[i+j*3,'Admin Factor - Lack of availability of blood transfusion'] = df[50][a+i+j*4]
            df_name.loc[i+j*3,'Admin Factor - Absence of trained stuff on duty'] = df[51][a+i+j*4]
            df_name.loc[i+j*3,'HW Factor - Delay in deciding to refer'] = df[53][a+i+j*4]
            df_name.loc[i+j*3,'HW Factor - Initial assessment incomplete'] = df[54][a+i+j*4]
            df_name.loc[i+j*3,'HW Factor - Inadequate resuscitation'] = df[55][a+i+j*4]
            df_name.loc[i+j*3,'HW Factor - Wrong diagnosis'] = df[56][a+i+j*4]
            df_name.loc[i+j*3,'HW Factor - Wrong treatment'] = df[57][a+i+j*4]
            df_name.loc[i+j*3,'HW Factor - No treatment'] = df[58][a+i+j*4]
            df_name.loc[i+j*3,'HW Factor - Delay in starting treatment'] = df[59][a+i+j*4]
            df_name.loc[i+j*3,'HW Factor - Inadequate monitoring'] = df[60][a+i+j*4]
            df_name.loc[i+j*3,'HW Factor - Prolonged abnormal observation without action'] = df[61][a+i+j*4]
            df_name.loc[i+j*3,'HW Factor - Lack of obstetric life saving skills'] = df[62][a+i+j*4]
            df_name.loc[i+j*3,'Patient Factor - Delay in reporting to health facility'] = df[64][a+i+j*4]
            df_name.loc[i+j*3,'Patient Factor - Lack of transport from home to health facility'] = df[65][a+i+j*4]
            df_name.loc[i+j*3,'Patient Factor - Unsafe traditional/cultural practice'] = df[66][a+i+j*4]
            df_name.loc[i+j*3,'Patient Factor - Refusal of treatment'] = df[67][a+i+j*4]
            df_name.loc[i+j*3,'Patient Factor - Delay in decision making'] = df[68][a+i+j*4]

    print(df_name)
