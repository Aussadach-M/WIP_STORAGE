import pandas as pd
import datetime as dt



def get_barcode_data(Path):
    try:
        Raw_Data = pd.read_excel(Path)
        Data_sheet =  pd.DataFrame()
        Data_sheet['Batch'] = Raw_Data['Batch']
        Data_sheet['GRTP'] = Raw_Data['Material Number (Material Description)']
        Data_sheet['SLOC'] = Raw_Data['Stor. Loc.']
        Data_sheet['Copyform'] = Raw_Data['Original Batch']
        Data_sheet['Cline'] = Raw_Data['Production Line']
        Data_sheet['Cdate'] = Raw_Data['Production Date']
        Data_sheet['Date_QC_ทำรับ'] = Raw_Data['Date of last goods receipt']
        Data_sheet['Weight'] = Raw_Data['Weight per Batch']
        Data_sheet['QC_Total'] = Raw_Data['QC-Total Decision']
        Data_sheet['BR_AQL'] = Raw_Data['BR-AQL']
        Data_sheet['CR_AQL'] = Raw_Data['CR-AQL']
        Data_sheet['MJ_AQL'] = Raw_Data['MJ-AQL']
        Data_sheet['MN_AQL'] = Raw_Data['MN-AQL']
        Data_sheet['PT_AQL'] = Raw_Data['PT-AQL']
        Data_sheet['Remark'] = Raw_Data['Remarks (QC1)']
        Data_sheet['Remark2'] = Raw_Data['Remarks (QC2)']
        Data_sheet['สถานะอายุ'] = ""
        Data_sheet['วันหมดอายุ_ครบ6เดือน'] = Raw_Data['Shelf Life Expiration Date']
        Data_sheet['ชิ้นพร้อมใช้UR'] = Raw_Data['Unrestricted']
        Data_sheet['Blocked'] = Raw_Data['Blocked']
        Data_sheet['รอตรวจสอบ'] = Raw_Data['Quality Insp.']
        return Data_sheet
    except:
        return pd.DataFrame()
    



def get_card_info(lookup_data=[],df=[]):
    # print(Product_inf)
    if(lookup_data != []):
        Df1 = pd.DataFrame(columns=['Way', 'Type', 'card','Time'])
        for i in lookup_data:
            if len(i)<4:

                Df1 = Df1.append({'Way': i[0],'Type':i[1],'card':i[2]},ignore_index=True)
            elif len(i)==4:

                Df1 = Df1.append({'Way': i[0],'Type':i[1],'card':i[2],'Time':i[3]},ignore_index=True)
        # with pd.option_context('display.max_rows', None, 'display.max_columns', None):  # more options can be specified also
        #     #print(df)
        #     print(Df1)
        pd.set_option("display.max_rows", None, "display.max_columns", None)
        #print(df)
    else: 
        print("No product provided")
        return None
    if df.empty:
        print("No product database provided")
        return None

    Left_join = pd.merge(Df1,  
                        df,  
                        left_on ='card',
                        right_on ='Batch',
                        how ='left') 
    pd.set_option("display.max_rows", None, "display.max_columns", None)

    
    Left_join.fillna('None', inplace=True)
    Row_list =[] 
    
    # Iterate over each row 
    for index, rows in Left_join.iterrows(): 
        # Create list for the current row 
        now = dt.datetime.now()
        time = now.strftime(r"%m-%d-%Y %H:%M:%S")
        update_time = time

        if rows.Way  in ('0',0):
            Status = "In storage"
        
        elif rows.Way in ('1',1):
            Status = "To Pack"

        elif rows.Way in ('2', 2) :
            Status = r"In storage(ReUpdate Data)"
            time = rows.Time
        else :
            Status = None
        #print("Size = {}".format(Left_join.loc[index, 'size']))

        my_list =[rows.Way, rows.Type, rows.card,rows.GRTP,rows.SLOC,rows.Copyform,rows.Cline,rows.Cdate,rows['Date_QC_ทำรับ'],rows.Weight,rows['QC_Total'],rows['BR_AQL'],rows['CR_AQL'],rows['MJ_AQL'],rows['MN_AQL'],rows['PT_AQL'],rows['Remark'],rows['Remark2'],rows['สถานะอายุ'],rows['วันหมดอายุ_ครบ6เดือน'],rows['ชิ้นพร้อมใช้UR'],rows['Blocked'],rows['รอตรวจสอบ'],Status,time,update_time] 
        # append the list to the final list 
        Row_list.append(my_list) 
    
    # Print the list 

    return Row_list


def get_card_info_for_update(lookup_data=[],df=[]):
    # print(Product_inf)
    if(lookup_data != []):
        Df1 = pd.DataFrame(columns=['Way', 'Type', 'card','Time'])
        for i in lookup_data:
            if len(i)<4:

                Df1 = Df1.append({'Way': i[0],'Type':i[1],'card':i[2]},ignore_index=True)
            elif len(i)==4:

                Df1 = Df1.append({'Way': i[0],'Type':i[1],'card':i[2],'Time':i[3]},ignore_index=True)
        # with pd.option_context('display.max_rows', None, 'display.max_columns', None):  # more options can be specified also
        #     #print(df)
        #     print(Df1)
        pd.set_option("display.max_rows", None, "display.max_columns", None)
        
    else: 
        print("No product provided")
        return None
    if df.empty:
        print("No product database provided")
        return None

    Left_join = pd.merge(Df1,  
                        df,  
                        left_on ='card',
                        right_on ='Batch',
                        how ='left') 


    #print(Left_join) 
    #Left_join = Left_join[Left_join['Batch'] >= 25]
    Left_join = Left_join[Left_join['Batch'].notna()]
    # Update only data that have value
    Left_join.fillna('None', inplace=True)
    
    Row_list =[] 
    
    # Iterate over each row 
    for index, rows in Left_join.iterrows(): 
        # Create list for the current row 
        now = dt.datetime.now()
        time = now.strftime(r"%m-%d-%Y %H:%M:%S")
        update_time = time

        if rows.Way  in ('0',0):
            Status = "In storage"
        
        elif rows.Way in ('1',1):
            Status = "To Pack"

        elif rows.Way in ('2', 2) :
            Status = r"In storage(ReUpdate Data)"
            time = rows.Time
        else :
            Status = None
        #print("Size = {}".format(Left_join.loc[index, 'size']))

        my_list =[rows.Way, rows.Type, rows.card,rows.GRTP,rows.SLOC,rows.Copyform,rows.Cline,rows.Cdate,rows['Date_QC_ทำรับ'],rows.Weight,rows['QC_Total'],rows['BR_AQL'],rows['CR_AQL'],rows['MJ_AQL'],rows['MN_AQL'],rows['PT_AQL'],rows['Remark'],rows['Remark2'],rows['สถานะอายุ'],rows['วันหมดอายุ_ครบ6เดือน'],rows['ชิ้นพร้อมใช้UR'],rows['Blocked'],rows['รอตรวจสอบ'],Status,time,update_time] 
        # append the list to the final list 
        Row_list.append(my_list) 
    
    # Print the list 

    return Row_list
