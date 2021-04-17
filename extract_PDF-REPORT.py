import pdfplumber
import os
import pandas as pd
import re

#Folder containing the PDFs
path = r'D:/PROJECTS_REPOSITORY/REPORTS/INPUT/'

#List of PDFs that could not be read
not_read = []

#DataFrame that saves the extracted information
df = pd.DataFrame([],columns=['RUC','EXERCISE_DETAIL','DATE_INFO_DETAIL','DATE_ISSUE','DETAIL','MONTH','QUANTITY','DATA_TYPE'])


for file in os.listdir(path):
    if 'pdf' in file:
        #File path
        fi_path = path+file
        
        try:
            #Reading the PDF
            pdf = pdfplumber.open(fi_path)
        except:
            continue
        
        #Extraction of RUC and DATE_ISSUE (they are always in the first page)
        ruc = pdf.pages[0].extract_text().split('RUC:')[1].split('\n')[0].strip()
        date_issue = pdf.pages[0].extract_text().split('Emitido el ')[1].split(' a')[0]
        
        try:
            for page in pdf.pages:
            
                if page.extract_text()!=None:
                    
                    #The table to be extracted contains as a title "INGRESOS NETOS DECLARADOS MENSUALMENTE"
                    if 'INGRESOS NETOS DECLARADOS MENSUALMENTE' in page.extract_text():
                        try:
                            num_page = page.page_number
                            
                            #This condition is used to extract the information in two scenarios, if the tables are on the same page or not
                            condition = 'no'
                            
                            #Extraction of EXERCISE_DETAIL and DATE_INFO_DETAIL
                            exercise_detail = 'EJERCICIO '+pdf.pages[num_page-1].extract_text().split('EJERCICIO')[1].split('\n')[0].strip().replace('- INGRESOS NETOS DECLARADOS MENSUALMENTE','').strip()
                            pre_text = pdf.pages[num_page-1].extract_text().split('INFORMACIÓN DE LAS DECLARACIONES MENSUALES')[1].split('- INGRESOS NETOS DECLARADOS')[0].replace('\n','').strip()
                            date_info_detail = re.findall(r"[0-9]+"r"/[0-9]+"r"/[0-9]+", pre_text)
                            if len(date_info_detail)==0:
                                date_info_detail = 'Información del '+re.findall(r"[0-9]"*4,exercise_detail)[0]
                            else:
                                date_info_detail = 'Información al '+date_info_detail[0]
                            
                            #Extracting all tables in current page
                            tablesList = pdf.pages[num_page-1].extract_tables()
                            
                            #Here I am extracting the information in the different scenarios found
                            if len(tablesList[-1][-1])>1:
                                
                                if tablesList[-1][-1][1]!='DICIEMBRE':
                                    
                                    df1 = pd.DataFrame(tablesList[-4],columns=['DETAIL','MONTH','QUANTITY','DATA_TYPE'])
                                    df2 = pd.DataFrame(tablesList[-2],columns=['DETAIL','MONTH','QUANTITY','DATA_TYPE'])
                                    
                                elif tablesList[-1][0][1]==tablesList[-2][0][1]:
                                    
                                    condition = 'yes'
                                    
                                    if 'EJERCICIO INMEDIATO ANTERIOR NO VENCIDO' not in pdf.pages[num_page-1].extract_text():
                                        page_text = pdf.pages[num_page-2].extract_text()
                                        exercise_detail = 'EJERCICIO '+page_text.split('EJERCICIO')[1].split('\n')[0].strip().replace('- INGRESOS NETOS DECLARADOS MENSUALMENTE','').strip()
                                        pre_text = page_text.split('INFORMACIÓN DE LAS DECLARACIONES MENSUALES')[1].split('- INGRESOS NETOS DECLARADOS')[0].replace('\n','').strip()
                                        date_info_detail = re.findall(r"[0-9]+"r"/[0-9]+"r"/[0-9]+", pre_text)
                                        if len(date_info_detail)==0:
                                            date_info_detail = 'Información del '+re.findall(r"[0-9]"*4,exercise_detail)[0]
                                        else:
                                            date_info_detail = 'Información al '+date_info_detail[0]
                                        
                                    exercise_detail2 = 'EJERCICIO ACTUAL '+pdf.pages[num_page-1].extract_text().split('EJERCICIO ACTUAL')[1].split('-')[0].strip().replace('- INGRESOS NETOS DECLARADOS MENSUALMENTE','').strip()
                                    pre_text2 = pdf.pages[num_page-1].extract_text().split('INFORMACIÓN DE LAS DECLARACIONES MENSUALES')[1]
                                    date_info_detail2 = re.findall(r"[0-9]+"r"/[0-9]+"r"/[0-9]+", pre_text2)
                                    if len(date_info_detail2)==0:
                                        date_info_detail2 = 'Información del '+re.findall(r"[0-9]"*4,exercise_detail2)[0]
                                    else:
                                        date_info_detail2 = 'Información al '+date_info_detail2[0]
                                    
                                    df1_1 = pd.DataFrame(tablesList[-4],columns=['DETAIL','MONTH','QUANTITY','DATA_TYPE'])
                                    df2_1 = pd.DataFrame(tablesList[-2],columns=['DETAIL','MONTH','QUANTITY','DATA_TYPE'])
                                    df1_2 = pd.DataFrame(tablesList[-3],columns=['DETAIL','MONTH','QUANTITY','DATA_TYPE'])
                                    df2_2 = pd.DataFrame(tablesList[-1],columns=['DETAIL','MONTH','QUANTITY','DATA_TYPE'])
                                    
                                    df1 = pd.concat([df1_1,df2_1])
                                    df1 =df1.reset_index(drop=True)
                                    #Creation of DataFrame that will contain the constant variables
                                    tam0 = df1.count()['DETAIL']
                                    li0 = []
                                    for ta in range(tam0):
                                        li0.append([ruc,exercise_detail,date_info_detail,date_issue])
                                    df_t = pd.DataFrame(li0,columns=['RUC','EXERCISE_DETAIL','DATE_INFO_DETAIL','DATE_ISSUE'])
                                    df1_temp = pd.concat([df_t,df1],axis=1)
                                    df1_temp.RUC.fillna(method='ffill',inplace=True)
                                    df1_temp.EXERCISE_DETAIL.fillna(method='ffill',inplace=True)
                                    df1_temp.DATE_INFO_DETAIL.fillna(method='ffill',inplace=True)
                                    df1_temp.DATE_ISSUE.fillna(method='ffill',inplace=True)
                                    df = pd.concat([df,df1_temp])
                                    
                                    df2 = pd.concat([df1_2,df2_2])
                                    df2 =df2.reset_index(drop=True)
                                    tam0 = df2.count()['DETAIL']
                                    li0 = []
                                    for ta in range(tam0):
                                        li0.append([ruc,exercise_detail2,date_info_detail2,date_issue])
                                    df_t = pd.DataFrame(li0,columns=['RUC','EXERCISE_DETAIL','DATE_INFO_DETAIL','DATE_ISSUE'])
                                    df2_temp = pd.concat([df_t,df2],axis=1)
                                    df2_temp.RUC.fillna(method='ffill',inplace=True)
                                    df2_temp.EXERCISE_DETAIL.fillna(method='ffill',inplace=True)
                                    df2_temp.DATE_INFO_DETAIL.fillna(method='ffill',inplace=True)
                                    df2_temp.DATE_ISSUE.fillna(method='ffill',inplace=True)
                                    df = pd.concat([df,df2_temp])
                                    
                                else:
                                    df1 = pd.DataFrame(tablesList[-2],columns=['DETAIL','MONTH','QUANTITY','DATA_TYPE'])
                                    df2 = pd.DataFrame(tablesList[-1],columns=['DETAIL','MONTH','QUANTITY','DATA_TYPE'])
    
                                if condition=='no':
                                    df_final = pd.concat([df1,df2])
                                    df_final["QUANTITY"] = df_final["QUANTITY"].astype('str').str.replace(',','')
                                    df_final =df_final.reset_index(drop=True)
                                    tam0 = df_final.count()['DETAIL']
                                    li0 = []
                                    for ta in range(tam0):
                                        li0.append([ruc,exercise_detail,date_info_detail,date_issue])
                                    df_t = pd.DataFrame(li0,columns=['RUC','EXERCISE_DETAIL','DATE_INFO_DETAIL','DATE_ISSUE'])
                                    df_temp = pd.concat([df_t,df_final],axis=1)
                                    df_temp.RUC.fillna(method='ffill',inplace=True)
                                    df_temp.EXERCISE_DETAIL.fillna(method='ffill',inplace=True)
                                    df_temp.DATE_INFO_DETAIL.fillna(method='ffill',inplace=True)
                                    df_temp.DATE_ISSUE.fillna(method='ffill',inplace=True)
                                    df = pd.concat([df,df_temp])
                                
                        except Exception as e:  
                            print(file)
                            print(e)
                            not_read.append([file])
                            
        except Exception as e:
            print(file)
            print(e)
            not_read.append([file])

#Save the DataFrame of the extracted data as an excel file           
df.to_excel(r'D:/PROJECTS_REPOSITORY/REPORTS/OUTPUT/MONTHLY_STATEMENTS.xlsx', sheet_name="MONTHLY_STATEMENTS", index=False)

#Creation and saving of DataFrame of unread PDF as excel file
dfNot_read = pd.DataFrame(not_read,columns=['FILE_NAME'])
dfNot_read.to_excel(r'D:/PROJECTS_REPOSITORY/REPORTS/OUTPUT/NOT_READ.xlsx', sheet_name="NOT_READ", index=False)
