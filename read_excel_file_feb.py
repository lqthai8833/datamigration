import openpyxl
from datetime import datetime, date
import time
from pathlib import Path
import logging
 
# Give the location of the file
 
xlsx_file = Path('E:\Working\DataMigration\Report\DailyPerformanceReport', '210714_Daily performance report_July_2021.xlsx')
wb_obj = openpyxl.load_workbook(xlsx_file) 

# Read the active sheet:
TTL_BRCH_CD = TTL_BRCH_NM = TTL_OPER_CD_B = TTL_OPER_CD_P = TOT_EMP_CNT = SE_TRD_FEE_B = SE_TRD_FEE_P = DE_TRD_VAL = TRD_DT = ''

sheet_list = wb_obj.sheetnames
sheet_list = ['0107', '0207', '0507', '0607', '0707', '0807', '0907', '1207', '1307', '1407']

for s in sheet_list:
    if (s != 'MS' and s != '2904'):

        insert_commands = []
        insert_sql_part01 = 'insert into RP01N002 (TRD_DT, TTL_BRCH_CD, TTL_BRCH_NM, TOT_EMP_CNT, TTL_OPER_CD, SE_TRD_FEE, DE_TRD_VAL, OPRT_IP_ADR, OPRT_DTTM) values ('

        sheet = wb_obj[s]
        TRD_DT = '2021' + s[2:4] + s[0:2]
        
        #sheet = wb_obj["0401"]
        #TRD_DT = '20210104'
        OPRT_DTTM = 'TO_DATE(\'' + str(datetime.now().strftime('%Y-%m-%d %H:%M:%S')) + '\', \'YYYY/MM/DD HH24:MI:SS\')'
        OPRT_IP_ADR = '0.0.0.0'
        

        #script_filename = 'dailyperformance_' + TRD_DT + '.sql'
        #print (script_filename)
        logging.basicConfig(filename = 'dailyperformance_' + TRD_DT + '.sql', filemode='w', format='%(levelname)s: %(message)s' ,level=logging.DEBUG)

        for i in range(2):
            if i == 0:
                if sheet["C16"].value == 'HQ: LM':
                    if "T1" in sheet["D16"].value:
                        TTL_BRCH_CD = 'BROKERHCM_5'
                        TTL_BRCH_NM = 'LM - TEAM 1'
                        bracket_pos_l = str(sheet["D16"].value).find('(')
                        bracket_pos_r = str(sheet["D16"].value).find(')')
                        TOT_EMP_CNT = (str(sheet["D16"].value))[bracket_pos_l+1:bracket_pos_r]
                        # type
                        TTL_OPER_CD_B = str(sheet["E16"].value).strip()
                        TTL_OPER_CD_P = str(sheet["E17"].value).strip()

                        # se daily value
                        SE_TRD_FEE_B = str(sheet["F16"].value)
                        SE_TRD_FEE_P = str(sheet["F17"].value)

                        #de daily value
                        DE_TRD_VAL_B = str(sheet["J16"].value)
                        DE_TRD_VAL_P = str(sheet["J17"].value)
                        if DE_TRD_VAL_B == 'None':
                            DE_TRD_VAL_B = '0'
                        if DE_TRD_VAL_P == 'None':
                            DE_TRD_VAL_P = '0'

                        # insert values for Broker
                        val_str = "', '".join([TRD_DT, TTL_BRCH_CD, TTL_BRCH_NM, TOT_EMP_CNT, TTL_OPER_CD_B, SE_TRD_FEE_B, DE_TRD_VAL_B, OPRT_IP_ADR])

                        if "=" in SE_TRD_FEE_B:
                            SE_TRD_FEE_B_REAL = SE_TRD_FEE_B[1:]
                            val_str = "', '".join([TRD_DT, TTL_BRCH_CD, TTL_BRCH_NM, TOT_EMP_CNT, TTL_OPER_CD_B])
                            val_str = val_str + "', " + SE_TRD_FEE_B_REAL + ", " + DE_TRD_VAL_B + ", '" + OPRT_IP_ADR 


                        insert_commands.append(insert_sql_part01 + '\'' + val_str + '\',' + OPRT_DTTM + ');')

                        # insert values for PR
                        val_str = "', '".join([TRD_DT, TTL_BRCH_CD, TTL_BRCH_NM, TOT_EMP_CNT, TTL_OPER_CD_P, SE_TRD_FEE_P, DE_TRD_VAL_P, OPRT_IP_ADR])

                        if "=" in SE_TRD_FEE_P:
                            SE_TRD_FEE_P_REAL = SE_TRD_FEE_P[1:]
                            val_str = "', '".join([TRD_DT, TTL_BRCH_CD, TTL_BRCH_NM, TOT_EMP_CNT, TTL_OPER_CD_P])
                            val_str = val_str + "', " + SE_TRD_FEE_P_REAL + ", " + DE_TRD_VAL_P + ", '" + OPRT_IP_ADR 

                        insert_commands.append(insert_sql_part01 + '\'' + val_str + '\',' + OPRT_DTTM + ');')

            else:
                if sheet["C16"].value == 'HQ: LM':
                    if "T2" in sheet["D19"].value:
                        TTL_BRCH_CD = 'LM_TEAM_2'
                        TTL_BRCH_NM = 'LM - TEAM 2'
                        bracket_pos_l = str(sheet["D19"].value).find('(')
                        bracket_pos_r = str(sheet["D19"].value).find(')')
                        TOT_EMP_CNT = (str(sheet["D19"].value))[bracket_pos_l+1:bracket_pos_r]
                        # type
                        TTL_OPER_CD_B = str(sheet["E19"].value).strip()
                        TTL_OPER_CD_P = str(sheet["E20"].value).strip()
                        # se daily value
                        SE_TRD_FEE_B = str(sheet["F19"].value)
                        SE_TRD_FEE_P = str(sheet["F20"].value)

                        #de daily value
                        DE_TRD_VAL_B = str(sheet["J19"].value)
                        DE_TRD_VAL_P = str(sheet["J20"].value)
                        if DE_TRD_VAL_B == 'None':
                            DE_TRD_VAL_B = '0'
                        if DE_TRD_VAL_P == 'None':
                            DE_TRD_VAL_P = '0'
                        

                        # insert values for Broker
                        val_str = "', '".join([TRD_DT, TTL_BRCH_CD, TTL_BRCH_NM, TOT_EMP_CNT, TTL_OPER_CD_B, SE_TRD_FEE_B, DE_TRD_VAL_B, OPRT_IP_ADR])
                        if "=" in SE_TRD_FEE_B:
                            SE_TRD_FEE_B_REAL = SE_TRD_FEE_B[1:]
                            val_str = "', '".join([TRD_DT, TTL_BRCH_CD, TTL_BRCH_NM, TOT_EMP_CNT, TTL_OPER_CD_B])
                            val_str = val_str + "', " + SE_TRD_FEE_B_REAL + ", " + DE_TRD_VAL_B + ", '" + OPRT_IP_ADR
                        
                        insert_commands.append(insert_sql_part01 + '\'' + val_str + '\',' + OPRT_DTTM + ');')

                        # insert values for PR
                        val_str = "', '".join([TRD_DT, TTL_BRCH_CD, TTL_BRCH_NM, TOT_EMP_CNT, TTL_OPER_CD_P, SE_TRD_FEE_P, DE_TRD_VAL_P, OPRT_IP_ADR])
                        if "=" in SE_TRD_FEE_P:
                            SE_TRD_FEE_P_REAL = SE_TRD_FEE_P[1:]
                            val_str = "', '".join([TRD_DT, TTL_BRCH_CD, TTL_BRCH_NM, TOT_EMP_CNT, TTL_OPER_CD_P])
                            val_str = val_str + "', " + SE_TRD_FEE_P_REAL + ", " + DE_TRD_VAL_P + ", '" + OPRT_IP_ADR 

                        insert_commands.append(insert_sql_part01 + '\'' + val_str + '\',' + OPRT_DTTM + ');')   

        # for SG branch
        for i in range(3):
            if i == 0:
                if sheet["C23"].value == 'SG':
                    if "T1" in sheet["D23"].value:
                        TTL_BRCH_CD = 'BROKERHCM_1'
                        TTL_BRCH_NM = 'SG - TEAM 1'
                        bracket_pos_l = str(sheet["D23"].value).find('(')
                        bracket_pos_r = str(sheet["D23"].value).find(')')

                        TOT_EMP_CNT = (str(sheet["D23"].value))[bracket_pos_l+1:bracket_pos_r]
                        # type
                        TTL_OPER_CD_B = str(sheet["E23"].value).strip()
                        TTL_OPER_CD_P = str(sheet["E24"].value).strip()
                        # se daily value
                        SE_TRD_FEE_B = str(sheet["F23"].value)
                        SE_TRD_FEE_P = str(sheet["F24"].value)

                        #de daily value
                        DE_TRD_VAL_B = str(sheet["J23"].value)
                        DE_TRD_VAL_P = str(sheet["J24"].value)
                        if DE_TRD_VAL_B == 'None':
                            DE_TRD_VAL_B = '0'
                        if DE_TRD_VAL_P == 'None':
                            DE_TRD_VAL_P = '0'

                        # insert values for Broker
                        val_str = "', '".join([TRD_DT, TTL_BRCH_CD, TTL_BRCH_NM, TOT_EMP_CNT, TTL_OPER_CD_B, SE_TRD_FEE_B, DE_TRD_VAL_B, OPRT_IP_ADR])
                        if "=" in SE_TRD_FEE_B:
                            SE_TRD_FEE_B_REAL = SE_TRD_FEE_B[1:]
                            val_str = "', '".join([TRD_DT, TTL_BRCH_CD, TTL_BRCH_NM, TOT_EMP_CNT, TTL_OPER_CD_B])
                            val_str = val_str + "', " + SE_TRD_FEE_B_REAL + ", " + DE_TRD_VAL_B + ", '" + OPRT_IP_ADR

                        insert_commands.append(insert_sql_part01 + '\'' + val_str + '\',' + OPRT_DTTM + ');')

                        # insert values for PR
                        val_str = "', '".join([TRD_DT, TTL_BRCH_CD, TTL_BRCH_NM, TOT_EMP_CNT, TTL_OPER_CD_P, SE_TRD_FEE_P, DE_TRD_VAL_P, OPRT_IP_ADR])
                        if "=" in SE_TRD_FEE_P:
                            SE_TRD_FEE_P_REAL = SE_TRD_FEE_P[1:]
                            val_str = "', '".join([TRD_DT, TTL_BRCH_CD, TTL_BRCH_NM, TOT_EMP_CNT, TTL_OPER_CD_P])
                            val_str = val_str + "', " + SE_TRD_FEE_P_REAL + ", " + DE_TRD_VAL_P + ", '" + OPRT_IP_ADR

                        insert_commands.append(insert_sql_part01 + '\'' + val_str + '\',' + OPRT_DTTM + ');') 

            elif i == 1:
                if "T2" in sheet["D26"].value:
                    TTL_BRCH_CD = 'BROKERHCM_3'
                    TTL_BRCH_NM = 'SG - TEAM 2'
                    bracket_pos_l = str(sheet["D26"].value).find('(')
                    bracket_pos_r = str(sheet["D26"].value).find(')')
                    TOT_EMP_CNT = (str(sheet["D26"].value))[bracket_pos_l+1:bracket_pos_r]
                    # type
                    TTL_OPER_CD_B = str(sheet["E26"].value).strip()
                    TTL_OPER_CD_P = str(sheet["E27"].value).strip()
                    # se daily value
                    SE_TRD_FEE_B = str(sheet["F26"].value)
                    SE_TRD_FEE_P = str(sheet["F27"].value)
                    
                    #de daily value
                    DE_TRD_VAL_B = str(sheet["J26"].value)
                    DE_TRD_VAL_P = str(sheet["J27"].value)
                    if DE_TRD_VAL_B == 'None':
                        DE_TRD_VAL_B = '0'
                    if DE_TRD_VAL_P == 'None':
                        DE_TRD_VAL_P = '0'

                    # insert values for Broker
                    val_str = "', '".join([TRD_DT, TTL_BRCH_CD, TTL_BRCH_NM, TOT_EMP_CNT, TTL_OPER_CD_B, SE_TRD_FEE_B, DE_TRD_VAL_B, OPRT_IP_ADR])
                    if "=" in SE_TRD_FEE_B:
                            SE_TRD_FEE_B_REAL = SE_TRD_FEE_B[1:]
                            val_str = "', '".join([TRD_DT, TTL_BRCH_CD, TTL_BRCH_NM, TOT_EMP_CNT, TTL_OPER_CD_B])
                            val_str = val_str + "', " + SE_TRD_FEE_B_REAL + ", " + DE_TRD_VAL_B + ", '" + OPRT_IP_ADR

                    insert_commands.append(insert_sql_part01 + '\'' + val_str + '\',' + OPRT_DTTM + ');')

                    # insert values for PR
                    val_str = "', '".join([TRD_DT, TTL_BRCH_CD, TTL_BRCH_NM, TOT_EMP_CNT, TTL_OPER_CD_P, SE_TRD_FEE_P, DE_TRD_VAL_P, OPRT_IP_ADR])
                    if "=" in SE_TRD_FEE_P:
                            SE_TRD_FEE_P_REAL = SE_TRD_FEE_P[1:]
                            val_str = "', '".join([TRD_DT, TTL_BRCH_CD, TTL_BRCH_NM, TOT_EMP_CNT, TTL_OPER_CD_P])
                            val_str = val_str + "', " + SE_TRD_FEE_P_REAL + ", " + DE_TRD_VAL_P + ", '" + OPRT_IP_ADR

                    insert_commands.append(insert_sql_part01 + '\'' + val_str + '\',' + OPRT_DTTM + ');')                   
            else:            
                if "T4" in sheet["D29"].value:
                    TTL_BRCH_CD = 'BROKERHCM_8'
                    TTL_BRCH_NM = 'SG - TEAM 4'
                    bracket_pos_l = str(sheet["D29"].value).find('(')
                    bracket_pos_r = str(sheet["D29"].value).find(')')
                    TOT_EMP_CNT = (str(sheet["D29"].value))[bracket_pos_l+1:bracket_pos_r]
                    # type
                    TTL_OPER_CD_B = str(sheet["E29"].value).strip()
                    TTL_OPER_CD_P = str(sheet["E30"].value).strip()
                    # se daily value
                    SE_TRD_FEE_B = str(sheet["F29"].value)
                    SE_TRD_FEE_P = str(sheet["F30"].value)
                    
                    #de daily value
                    DE_TRD_VAL_B = str(sheet["J29"].value)
                    DE_TRD_VAL_P = str(sheet["J30"].value)
                    if DE_TRD_VAL_B == 'None':
                        DE_TRD_VAL_B = '0'
                    if DE_TRD_VAL_P == 'None':
                        DE_TRD_VAL_P = '0'

                    # insert values for Broker
                    val_str = "', '".join([TRD_DT, TTL_BRCH_CD, TTL_BRCH_NM, TOT_EMP_CNT, TTL_OPER_CD_B, SE_TRD_FEE_B, DE_TRD_VAL_B, OPRT_IP_ADR])
                    if "=" in SE_TRD_FEE_B:
                            SE_TRD_FEE_B_REAL = SE_TRD_FEE_B[1:]
                            val_str = "', '".join([TRD_DT, TTL_BRCH_CD, TTL_BRCH_NM, TOT_EMP_CNT, TTL_OPER_CD_B])
                            val_str = val_str + "', " + SE_TRD_FEE_B_REAL + ", " + DE_TRD_VAL_B + ", '" + OPRT_IP_ADR

                    insert_commands.append(insert_sql_part01 + '\'' + val_str + '\',' + OPRT_DTTM + ');')

                    # insert values for PR
                    val_str = "', '".join([TRD_DT, TTL_BRCH_CD, TTL_BRCH_NM, TOT_EMP_CNT, TTL_OPER_CD_P, SE_TRD_FEE_P, DE_TRD_VAL_P, OPRT_IP_ADR])
                    if "=" in SE_TRD_FEE_P:
                            SE_TRD_FEE_P_REAL = SE_TRD_FEE_P[1:]
                            val_str = "', '".join([TRD_DT, TTL_BRCH_CD, TTL_BRCH_NM, TOT_EMP_CNT, TTL_OPER_CD_P])
                            val_str = val_str + "', " + SE_TRD_FEE_P_REAL + ", " + DE_TRD_VAL_P + ", '" + OPRT_IP_ADR

                    insert_commands.append(insert_sql_part01 + '\'' + val_str + '\',' + OPRT_DTTM + ');') 
        # New branch
        if "CMT8" in sheet["C33"].value:
            TTL_BRCH_CD = 'BROKERHCM_9'
            TTL_BRCH_NM = 'CMT8'
            bracket_pos_l = str(sheet["C33"].value).find('(')
            bracket_pos_r = str(sheet["C33"].value).find(')')
            TOT_EMP_CNT = (str(sheet["C33"].value))[bracket_pos_l+1:bracket_pos_r]
            # type
            TTL_OPER_CD_B = str(sheet["E33"].value).strip()
            TTL_OPER_CD_P = str(sheet["E34"].value).strip()
            # se daily value
            SE_TRD_FEE_B = str(sheet["F33"].value)
            SE_TRD_FEE_P = str(sheet["F34"].value)

            #de daily value
            DE_TRD_VAL_B = str(sheet["J33"].value)
            DE_TRD_VAL_P = str(sheet["J34"].value)
            if DE_TRD_VAL_B == 'None':
                DE_TRD_VAL_B = '0'
            if DE_TRD_VAL_P == 'None':
                DE_TRD_VAL_P = '0'

            # insert values for Broker
            val_str = "', '".join([TRD_DT, TTL_BRCH_CD, TTL_BRCH_NM, TOT_EMP_CNT, TTL_OPER_CD_B, SE_TRD_FEE_B, DE_TRD_VAL_B, OPRT_IP_ADR])
            if "=" in SE_TRD_FEE_B:
                            SE_TRD_FEE_B_REAL = SE_TRD_FEE_B[1:]
                            val_str = "', '".join([TRD_DT, TTL_BRCH_CD, TTL_BRCH_NM, TOT_EMP_CNT, TTL_OPER_CD_B])
                            val_str = val_str + "', " + SE_TRD_FEE_B_REAL + ", " + DE_TRD_VAL_B + ", '" + OPRT_IP_ADR

            insert_commands.append(insert_sql_part01 + '\'' + val_str + '\',' + OPRT_DTTM + ');')

            # insert values for PR
            val_str = "', '".join([TRD_DT, TTL_BRCH_CD, TTL_BRCH_NM, TOT_EMP_CNT, TTL_OPER_CD_P, SE_TRD_FEE_P, DE_TRD_VAL_P, OPRT_IP_ADR])
            if "=" in SE_TRD_FEE_P:
                            SE_TRD_FEE_P_REAL = SE_TRD_FEE_P[1:]
                            val_str = "', '".join([TRD_DT, TTL_BRCH_CD, TTL_BRCH_NM, TOT_EMP_CNT, TTL_OPER_CD_P])
                            val_str = val_str + "', " + SE_TRD_FEE_P_REAL + ", " + DE_TRD_VAL_P + ", '" + OPRT_IP_ADR

            insert_commands.append(insert_sql_part01 + '\'' + val_str + '\',' + OPRT_DTTM + ');')

        # HCM branch
        if "HCM" in sheet["C36"].value:
            TTL_BRCH_CD = 'BROKERHCM_4'
            TTL_BRCH_NM = 'HCM'
            bracket_pos_l = str(sheet["C36"].value).find('(')
            bracket_pos_r = str(sheet["C36"].value).find(')')
            TOT_EMP_CNT = (str(sheet["C36"].value))[bracket_pos_l+1:bracket_pos_r]
            # type
            TTL_OPER_CD_B = str(sheet["E36"].value).strip()
            TTL_OPER_CD_P = str(sheet["E37"].value).strip()
            # se daily value
            SE_TRD_FEE_B = str(sheet["F36"].value)
            SE_TRD_FEE_P = str(sheet["F37"].value)

            #de daily value
            DE_TRD_VAL_B = str(sheet["J36"].value)
            DE_TRD_VAL_P = str(sheet["J37"].value)
            if DE_TRD_VAL_B == 'None':
                DE_TRD_VAL_B = '0'
            if DE_TRD_VAL_P == 'None':
                DE_TRD_VAL_P = '0'

            # insert values for Broker
            val_str = "', '".join([TRD_DT, TTL_BRCH_CD, TTL_BRCH_NM, TOT_EMP_CNT, TTL_OPER_CD_B, SE_TRD_FEE_B, DE_TRD_VAL_B, OPRT_IP_ADR])
            if "=" in SE_TRD_FEE_B:
                            SE_TRD_FEE_B_REAL = SE_TRD_FEE_B[1:]
                            val_str = "', '".join([TRD_DT, TTL_BRCH_CD, TTL_BRCH_NM, TOT_EMP_CNT, TTL_OPER_CD_B])
                            val_str = val_str + "', " + SE_TRD_FEE_B_REAL + ", " + DE_TRD_VAL_B + ", '" + OPRT_IP_ADR

            insert_commands.append(insert_sql_part01 + '\'' + val_str + '\',' + OPRT_DTTM + ');')

            # insert values for PR
            val_str = "', '".join([TRD_DT, TTL_BRCH_CD, TTL_BRCH_NM, TOT_EMP_CNT, TTL_OPER_CD_P, SE_TRD_FEE_P, DE_TRD_VAL_P, OPRT_IP_ADR])
            if "=" in SE_TRD_FEE_P:
                            SE_TRD_FEE_P_REAL = SE_TRD_FEE_P[1:]
                            val_str = "', '".join([TRD_DT, TTL_BRCH_CD, TTL_BRCH_NM, TOT_EMP_CNT, TTL_OPER_CD_P])
                            val_str = val_str + "', " + SE_TRD_FEE_P_REAL + ", " + DE_TRD_VAL_P + ", '" + OPRT_IP_ADR

            insert_commands.append(insert_sql_part01 + '\'' + val_str + '\',' + OPRT_DTTM + ');')

        # DN branch
        if "DN" in sheet["C39"].value:
            TTL_BRCH_CD = 'BROKERDN'
            TTL_BRCH_NM = 'DN'
            bracket_pos_l = str(sheet["C39"].value).find('(')
            bracket_pos_r = str(sheet["C39"].value).find(')')
            TOT_EMP_CNT = (str(sheet["C39"].value))[bracket_pos_l+1:bracket_pos_r]
            # type
            TTL_OPER_CD_B = str(sheet["E39"].value).strip()
            TTL_OPER_CD_P = str(sheet["E40"].value).strip()
            # se daily value
            SE_TRD_FEE_B = str(sheet["F39"].value)
            SE_TRD_FEE_P = str(sheet["F40"].value)

            #de daily value
            DE_TRD_VAL_B = str(sheet["J39"].value)
            DE_TRD_VAL_P = str(sheet["J40"].value)
            if DE_TRD_VAL_B == 'None':
                DE_TRD_VAL_B = '0'
            if DE_TRD_VAL_P == 'None':
                DE_TRD_VAL_P = '0'

            # insert values for Broker
            val_str = "', '".join([TRD_DT, TTL_BRCH_CD, TTL_BRCH_NM, TOT_EMP_CNT, TTL_OPER_CD_B, SE_TRD_FEE_B, DE_TRD_VAL_B, OPRT_IP_ADR])
            if "=" in SE_TRD_FEE_B:
                            SE_TRD_FEE_B_REAL = SE_TRD_FEE_B[1:]
                            val_str = "', '".join([TRD_DT, TTL_BRCH_CD, TTL_BRCH_NM, TOT_EMP_CNT, TTL_OPER_CD_B])
                            val_str = val_str + "', " + SE_TRD_FEE_B_REAL + ", " + DE_TRD_VAL_B + ", '" + OPRT_IP_ADR

            insert_commands.append(insert_sql_part01 + '\'' + val_str + '\',' + OPRT_DTTM + ');')

            # insert values for PR
            val_str = "', '".join([TRD_DT, TTL_BRCH_CD, TTL_BRCH_NM, TOT_EMP_CNT, TTL_OPER_CD_P, SE_TRD_FEE_P, DE_TRD_VAL_P, OPRT_IP_ADR])
            if "=" in SE_TRD_FEE_P:
                            SE_TRD_FEE_P_REAL = SE_TRD_FEE_P[1:]
                            val_str = "', '".join([TRD_DT, TTL_BRCH_CD, TTL_BRCH_NM, TOT_EMP_CNT, TTL_OPER_CD_P])
                            val_str = val_str + "', " + SE_TRD_FEE_P_REAL + ", " + DE_TRD_VAL_P + ", '" + OPRT_IP_ADR

            insert_commands.append(insert_sql_part01 + '\'' + val_str + '\',' + OPRT_DTTM + ');')

        # DN branch
        if "VT" in sheet["C42"].value:
            TTL_BRCH_CD = 'BROKERVT'
            TTL_BRCH_NM = 'VT'
            bracket_pos_l = str(sheet["C42"].value).find('(')
            bracket_pos_r = str(sheet["C42"].value).find(')')
            TOT_EMP_CNT = (str(sheet["C42"].value))[bracket_pos_l+1:bracket_pos_r]
            # type
            TTL_OPER_CD_B = str(sheet["E42"].value).strip()
            TTL_OPER_CD_P = str(sheet["E43"].value).strip()
            # se daily value
            SE_TRD_FEE_B = str(sheet["F42"].value)
            SE_TRD_FEE_P = str(sheet["F43"].value)

            #de daily value
            DE_TRD_VAL_B = str(sheet["J42"].value)
            DE_TRD_VAL_P = str(sheet["J43"].value)
            if DE_TRD_VAL_B == 'None':
                DE_TRD_VAL_B = '0'
            if DE_TRD_VAL_P == 'None':
                DE_TRD_VAL_P = '0'

            # insert values for Broker
            val_str = "', '".join([TRD_DT, TTL_BRCH_CD, TTL_BRCH_NM, TOT_EMP_CNT, TTL_OPER_CD_B, SE_TRD_FEE_B, DE_TRD_VAL_B, OPRT_IP_ADR])
            if "=" in SE_TRD_FEE_B:
                            SE_TRD_FEE_B_REAL = SE_TRD_FEE_B[1:]
                            val_str = "', '".join([TRD_DT, TTL_BRCH_CD, TTL_BRCH_NM, TOT_EMP_CNT, TTL_OPER_CD_B])
                            val_str = val_str + "', " + SE_TRD_FEE_B_REAL + ", " + DE_TRD_VAL_B + ", '" + OPRT_IP_ADR
            insert_commands.append(insert_sql_part01 + '\'' + val_str + '\',' + OPRT_DTTM + ');')

            # insert values for PR
            val_str = "', '".join([TRD_DT, TTL_BRCH_CD, TTL_BRCH_NM, TOT_EMP_CNT, TTL_OPER_CD_P, SE_TRD_FEE_P, DE_TRD_VAL_P, OPRT_IP_ADR])
            if "=" in SE_TRD_FEE_P:
                            SE_TRD_FEE_P_REAL = SE_TRD_FEE_P[1:]
                            val_str = "', '".join([TRD_DT, TTL_BRCH_CD, TTL_BRCH_NM, TOT_EMP_CNT, TTL_OPER_CD_P])
                            val_str = val_str + "', " + SE_TRD_FEE_P_REAL + ", " + DE_TRD_VAL_P + ", '" + OPRT_IP_ADR
            insert_commands.append(insert_sql_part01 + '\'' + val_str + '\',' + OPRT_DTTM + ');')

        if "CT" in sheet["C45"].value:
            TTL_BRCH_CD = 'BROKERCT'
            TTL_BRCH_NM = 'CT'
            bracket_pos_l = str(sheet["C45"].value).find('(')
            bracket_pos_r = str(sheet["C45"].value).find(')')
            TOT_EMP_CNT = (str(sheet["C45"].value))[bracket_pos_l+1:bracket_pos_r]
            # type
            TTL_OPER_CD_B = str(sheet["E45"].value).strip()
            TTL_OPER_CD_P = str(sheet["E46"].value).strip()
            # se daily value
            SE_TRD_FEE_B = str(sheet["F45"].value)
            SE_TRD_FEE_P = str(sheet["F46"].value)
            #de daily value
            DE_TRD_VAL_B = str(sheet["J45"].value)
            DE_TRD_VAL_P = str(sheet["J46"].value)
            if DE_TRD_VAL_B == 'None':
                DE_TRD_VAL_B = '0'
            if DE_TRD_VAL_P == 'None':
                DE_TRD_VAL_P = '0'

            # insert values for Broker
            val_str = "', '".join([TRD_DT, TTL_BRCH_CD, TTL_BRCH_NM, TOT_EMP_CNT, TTL_OPER_CD_B, SE_TRD_FEE_B, DE_TRD_VAL_B, OPRT_IP_ADR])
            if "=" in SE_TRD_FEE_B:
                            SE_TRD_FEE_B_REAL = SE_TRD_FEE_B[1:]
                            val_str = "', '".join([TRD_DT, TTL_BRCH_CD, TTL_BRCH_NM, TOT_EMP_CNT, TTL_OPER_CD_B])
                            val_str = val_str + "', " + SE_TRD_FEE_B_REAL + ", " + DE_TRD_VAL_B + ", '" + OPRT_IP_ADR
            insert_commands.append(insert_sql_part01 + '\'' + val_str + '\',' + OPRT_DTTM + ');')

            # insert values for PR
            val_str = "', '".join([TRD_DT, TTL_BRCH_CD, TTL_BRCH_NM, TOT_EMP_CNT, TTL_OPER_CD_P, SE_TRD_FEE_P, DE_TRD_VAL_P, OPRT_IP_ADR])
            if "=" in SE_TRD_FEE_P:
                            SE_TRD_FEE_P_REAL = SE_TRD_FEE_P[1:]
                            val_str = "', '".join([TRD_DT, TTL_BRCH_CD, TTL_BRCH_NM, TOT_EMP_CNT, TTL_OPER_CD_P])
                            val_str = val_str + "', " + SE_TRD_FEE_P_REAL + ", " + DE_TRD_VAL_P + ", '" + OPRT_IP_ADR
            insert_commands.append(insert_sql_part01 + '\'' + val_str + '\',' + OPRT_DTTM + ');')    

        if "HN" in sheet["C48"].value:
            TTL_BRCH_CD = 'BROKERHN'
            TTL_BRCH_NM = 'HN'
            bracket_pos_l = str(sheet["C48"].value).find('(')
            bracket_pos_r = str(sheet["C48"].value).find(')')
            TOT_EMP_CNT = (str(sheet["C48"].value))[bracket_pos_l+1:bracket_pos_r]
            # type
            TTL_OPER_CD_B = str(sheet["E48"].value).strip()
            TTL_OPER_CD_P = str(sheet["E49"].value).strip()
            # se daily value
            SE_TRD_FEE_B = str(sheet["F48"].value)
            SE_TRD_FEE_P = str(sheet["F49"].value)
            #de daily value
            DE_TRD_VAL_B = str(sheet["J48"].value)
            DE_TRD_VAL_P = str(sheet["J49"].value)
            if DE_TRD_VAL_B == 'None':
                DE_TRD_VAL_B = '0'
            if DE_TRD_VAL_P == 'None':
                DE_TRD_VAL_P = '0'

            # insert values for Broker
            val_str = "', '".join([TRD_DT, TTL_BRCH_CD, TTL_BRCH_NM, TOT_EMP_CNT, TTL_OPER_CD_B, SE_TRD_FEE_B, DE_TRD_VAL_B, OPRT_IP_ADR])
            if "=" in SE_TRD_FEE_B:
                            SE_TRD_FEE_B_REAL = SE_TRD_FEE_B[1:]
                            val_str = "', '".join([TRD_DT, TTL_BRCH_CD, TTL_BRCH_NM, TOT_EMP_CNT, TTL_OPER_CD_B])
                            val_str = val_str + "', " + SE_TRD_FEE_B_REAL + ", " + DE_TRD_VAL_B + ", '" + OPRT_IP_ADR
            insert_commands.append(insert_sql_part01 + '\'' + val_str + '\',' + OPRT_DTTM + ');')

            # insert values for PR
            val_str = "', '".join([TRD_DT, TTL_BRCH_CD, TTL_BRCH_NM, TOT_EMP_CNT, TTL_OPER_CD_P, SE_TRD_FEE_P, DE_TRD_VAL_P, OPRT_IP_ADR])
            if "=" in SE_TRD_FEE_P:
                            SE_TRD_FEE_P_REAL = SE_TRD_FEE_P[1:]
                            val_str = "', '".join([TRD_DT, TTL_BRCH_CD, TTL_BRCH_NM, TOT_EMP_CNT, TTL_OPER_CD_P])
                            val_str = val_str + "', " + SE_TRD_FEE_P_REAL + ", " + DE_TRD_VAL_P + ", '" + OPRT_IP_ADR
            insert_commands.append(insert_sql_part01 + '\'' + val_str + '\',' + OPRT_DTTM + ');')      

        if "HN Biz Center" in sheet["C51"].value:
            TTL_BRCH_CD = 'BROKERHN1'
            TTL_BRCH_NM = 'HN Biz Center'
            bracket_pos_l = str(sheet["C51"].value).find('(')
            bracket_pos_r = str(sheet["C51"].value).find(')')
            TOT_EMP_CNT = (str(sheet["C51"].value))[bracket_pos_l+1:bracket_pos_r]
            # type
            TTL_OPER_CD_B = str(sheet["E51"].value).strip()
            TTL_OPER_CD_P = str(sheet["E52"].value).strip()
            # se daily value
            SE_TRD_FEE_B = str(sheet["F51"].value)
            SE_TRD_FEE_P = str(sheet["F52"].value)

            #de daily value
            DE_TRD_VAL_B = str(sheet["J51"].value)
            DE_TRD_VAL_P = str(sheet["J52"].value)
            if DE_TRD_VAL_B == 'None':
                DE_TRD_VAL_B = '0'
            if DE_TRD_VAL_P == 'None':
                DE_TRD_VAL_P = '0'

            # insert values for Broker
            val_str = "', '".join([TRD_DT, TTL_BRCH_CD, TTL_BRCH_NM, TOT_EMP_CNT, TTL_OPER_CD_B, SE_TRD_FEE_B, DE_TRD_VAL_B, OPRT_IP_ADR])
            if "=" in SE_TRD_FEE_B:
                            SE_TRD_FEE_B_REAL = SE_TRD_FEE_B[1:]
                            val_str = "', '".join([TRD_DT, TTL_BRCH_CD, TTL_BRCH_NM, TOT_EMP_CNT, TTL_OPER_CD_B])
                            val_str = val_str + "', " + SE_TRD_FEE_B_REAL + ", " + DE_TRD_VAL_B + ", '" + OPRT_IP_ADR
            insert_commands.append(insert_sql_part01 + '\'' + val_str + '\',' + OPRT_DTTM + ');')

            # insert values for PR
            val_str = "', '".join([TRD_DT, TTL_BRCH_CD, TTL_BRCH_NM, TOT_EMP_CNT, TTL_OPER_CD_P, SE_TRD_FEE_P, DE_TRD_VAL_P, OPRT_IP_ADR])
            if "=" in SE_TRD_FEE_P:
                            SE_TRD_FEE_P_REAL = SE_TRD_FEE_P[1:]
                            val_str = "', '".join([TRD_DT, TTL_BRCH_CD, TTL_BRCH_NM, TOT_EMP_CNT, TTL_OPER_CD_P])
                            val_str = val_str + "', " + SE_TRD_FEE_P_REAL + ", " + DE_TRD_VAL_P + ", '" + OPRT_IP_ADR
            insert_commands.append(insert_sql_part01 + '\'' + val_str + '\',' + OPRT_DTTM + ');') 

        if "ETF Trading" in sheet["C54"].value:
            TTL_BRCH_CD = 'ETF Trading'
            TTL_BRCH_NM = 'ETF Trading'
            bracket_pos_l = str(sheet["C54"].value).find('(')
            bracket_pos_r = str(sheet["C54"].value).find(')')
            TOT_EMP_CNT = (str(sheet["C54"].value))[bracket_pos_l+1:bracket_pos_r]
            # type
            TTL_OPER_CD_B = str(sheet["E54"].value).strip()
            TTL_OPER_CD_P = str(sheet["E55"].value).strip()
            # se daily value
            SE_TRD_FEE_B = str(sheet["F54"].value)
            SE_TRD_FEE_P = str(sheet["F55"].value)
            if SE_TRD_FEE_B == 'None':
                SE_TRD_FEE_B = '0'
            if SE_TRD_FEE_P == 'None':
                SE_TRD_FEE_P = '0'

            #de daily value
            DE_TRD_VAL_B = str(sheet["J54"].value)
            DE_TRD_VAL_P = str(sheet["J55"].value)
            if DE_TRD_VAL_B == 'None':
                DE_TRD_VAL_B = '0'
            if DE_TRD_VAL_P == 'None':
                DE_TRD_VAL_P = '0'

            # insert values for Broker
            val_str = "', '".join([TRD_DT, TTL_BRCH_CD, TTL_BRCH_NM, TOT_EMP_CNT, TTL_OPER_CD_B, SE_TRD_FEE_B, DE_TRD_VAL_B, OPRT_IP_ADR])
            if "=" in SE_TRD_FEE_B:
                            SE_TRD_FEE_B_REAL = SE_TRD_FEE_B[1:]
                            val_str = "', '".join([TRD_DT, TTL_BRCH_CD, TTL_BRCH_NM, TOT_EMP_CNT, TTL_OPER_CD_B])
                            val_str = val_str + "', " + SE_TRD_FEE_B_REAL + ", " + DE_TRD_VAL_B + ", '" + OPRT_IP_ADR
            insert_commands.append(insert_sql_part01 + '\'' + val_str + '\',' + OPRT_DTTM + ');')

            # insert values for PR
            val_str = "', '".join([TRD_DT, TTL_BRCH_CD, TTL_BRCH_NM, TOT_EMP_CNT, TTL_OPER_CD_P, SE_TRD_FEE_P, DE_TRD_VAL_P, OPRT_IP_ADR])
            if "=" in SE_TRD_FEE_P:
                            SE_TRD_FEE_P_REAL = SE_TRD_FEE_P[1:]
                            val_str = "', '".join([TRD_DT, TTL_BRCH_CD, TTL_BRCH_NM, TOT_EMP_CNT, TTL_OPER_CD_P])
                            val_str = val_str + "', " + SE_TRD_FEE_P_REAL + ", " + DE_TRD_VAL_P + ", '" + OPRT_IP_ADR
            insert_commands.append(insert_sql_part01 + '\'' + val_str + '\',' + OPRT_DTTM + ');') 

        if "TL" in sheet["C57"].value:
            TTL_BRCH_CD = 'BROKERHN2'
            TTL_BRCH_NM = 'TL'
            bracket_pos_l = str(sheet["C57"].value).find('(')
            bracket_pos_r = str(sheet["C57"].value).find(')')
            TOT_EMP_CNT = (str(sheet["C57"].value))[bracket_pos_l+1:bracket_pos_r]
            # type
            TTL_OPER_CD_B = str(sheet["E57"].value).strip()
            TTL_OPER_CD_P = str(sheet["E58"].value).strip()
            # se daily value
            SE_TRD_FEE_B = str(sheet["F57"].value)
            SE_TRD_FEE_P = str(sheet["F58"].value)

            #de daily value
            DE_TRD_VAL_B = str(sheet["J57"].value)
            DE_TRD_VAL_P = str(sheet["J58"].value)
            if DE_TRD_VAL_B == 'None':
                DE_TRD_VAL_B = '0'
            if DE_TRD_VAL_P == 'None':
                DE_TRD_VAL_P = '0'

            # insert values for Broker
            val_str = "', '".join([TRD_DT, TTL_BRCH_CD, TTL_BRCH_NM, TOT_EMP_CNT, TTL_OPER_CD_B, SE_TRD_FEE_B, DE_TRD_VAL_B, OPRT_IP_ADR])
            if "=" in SE_TRD_FEE_B:
                            SE_TRD_FEE_B_REAL = SE_TRD_FEE_B[1:]
                            val_str = "', '".join([TRD_DT, TTL_BRCH_CD, TTL_BRCH_NM, TOT_EMP_CNT, TTL_OPER_CD_B])
                            val_str = val_str + "', " + SE_TRD_FEE_B_REAL + ", " + DE_TRD_VAL_B + ", '" + OPRT_IP_ADR
            insert_commands.append(insert_sql_part01 + '\'' + val_str + '\',' + OPRT_DTTM + ');')

            # insert values for PR
            val_str = "', '".join([TRD_DT, TTL_BRCH_CD, TTL_BRCH_NM, TOT_EMP_CNT, TTL_OPER_CD_P, SE_TRD_FEE_P, DE_TRD_VAL_P, OPRT_IP_ADR])
            if "=" in SE_TRD_FEE_P:
                            SE_TRD_FEE_P_REAL = SE_TRD_FEE_P[1:]
                            val_str = "', '".join([TRD_DT, TTL_BRCH_CD, TTL_BRCH_NM, TOT_EMP_CNT, TTL_OPER_CD_P])
                            val_str = val_str + "', " + SE_TRD_FEE_P_REAL + ", " + DE_TRD_VAL_P + ", '" + OPRT_IP_ADR
            insert_commands.append(insert_sql_part01 + '\'' + val_str + '\',' + OPRT_DTTM + ');') 

        if "HP" in sheet["C60"].value:
            TTL_BRCH_CD = 'BRKHP1_1'
            TTL_BRCH_NM = 'HP'
            bracket_pos_l = str(sheet["C60"].value).find('(')
            bracket_pos_r = str(sheet["C60"].value).find(')')
            TOT_EMP_CNT = (str(sheet["C60"].value))[bracket_pos_l+1:bracket_pos_r]
            # type
            TTL_OPER_CD_B = str(sheet["E60"].value).strip()
            TTL_OPER_CD_P = str(sheet["E61"].value).strip()
            # se daily value
            SE_TRD_FEE_B = str(sheet["F60"].value)
            SE_TRD_FEE_P = str(sheet["F61"].value)

            #de daily value
            DE_TRD_VAL_B = str(sheet["J60"].value)
            DE_TRD_VAL_P = str(sheet["J61"].value)
            if DE_TRD_VAL_B == 'None':
                DE_TRD_VAL_B = '0'
            if DE_TRD_VAL_P == 'None':
                DE_TRD_VAL_P = '0'

            # insert values for Broker
            val_str = "', '".join([TRD_DT, TTL_BRCH_CD, TTL_BRCH_NM, TOT_EMP_CNT, TTL_OPER_CD_B, SE_TRD_FEE_B, DE_TRD_VAL_B, OPRT_IP_ADR])

            if "=" in SE_TRD_FEE_B:
                SE_TRD_FEE_B_REAL = SE_TRD_FEE_B[1:]
                val_str = "', '".join([TRD_DT, TTL_BRCH_CD, TTL_BRCH_NM, TOT_EMP_CNT, TTL_OPER_CD_B])
                val_str = val_str + "', " + SE_TRD_FEE_B_REAL + ", " + DE_TRD_VAL_B + ", '" + OPRT_IP_ADR 

            insert_commands.append(insert_sql_part01 + '\'' + val_str + '\',' + OPRT_DTTM + ');')

            # insert values for PR
            val_str = "', '".join([TRD_DT, TTL_BRCH_CD, TTL_BRCH_NM, TOT_EMP_CNT, TTL_OPER_CD_P, SE_TRD_FEE_P, DE_TRD_VAL_P, OPRT_IP_ADR])

            if "=" in SE_TRD_FEE_P:
                SE_TRD_FEE_P_REAL = SE_TRD_FEE_P[1:]
                val_str = "', '".join([TRD_DT, TTL_BRCH_CD, TTL_BRCH_NM, TOT_EMP_CNT, TTL_OPER_CD_P])
                val_str = val_str + "', " + SE_TRD_FEE_P_REAL + ", " + DE_TRD_VAL_P + ", '" + OPRT_IP_ADR 
            
            insert_commands.append(insert_sql_part01 + '\'' + val_str + '\',' + OPRT_DTTM + ');')     

        if "Retail 1" in sheet["C63"].value:
            TTL_BRCH_CD = 'BROKERHCM'
            TTL_BRCH_NM = 'Retail 1'
            bracket_pos_l = str(sheet["C63"].value).find('(')
            bracket_pos_r = str(sheet["C63"].value).find(')')
            TOT_EMP_CNT = (str(sheet["C63"].value))[bracket_pos_l+1:bracket_pos_r]
            # type
            TTL_OPER_CD_B = str(sheet["E63"].value).strip()
            TTL_OPER_CD_P = str(sheet["E64"].value).strip()
            # se daily value
            SE_TRD_FEE_B = str(sheet["F63"].value)
            SE_TRD_FEE_P = str(sheet["F64"].value)

            #de daily value
            DE_TRD_VAL_B = str(sheet["J63"].value)
            DE_TRD_VAL_P = str(sheet["J64"].value)
            if DE_TRD_VAL_B == 'None':
                DE_TRD_VAL_B = '0'
            if DE_TRD_VAL_P == 'None':
                DE_TRD_VAL_P = '0'

            # insert values for Broker
            val_str = "', '".join([TRD_DT, TTL_BRCH_CD, TTL_BRCH_NM, TOT_EMP_CNT, TTL_OPER_CD_B, SE_TRD_FEE_B, DE_TRD_VAL_B, OPRT_IP_ADR])
            if "=" in SE_TRD_FEE_B:
                            SE_TRD_FEE_B_REAL = SE_TRD_FEE_B[1:]
                            val_str = "', '".join([TRD_DT, TTL_BRCH_CD, TTL_BRCH_NM, TOT_EMP_CNT, TTL_OPER_CD_B])
                            val_str = val_str + "', " + SE_TRD_FEE_B_REAL + ", " + DE_TRD_VAL_B + ", '" + OPRT_IP_ADR
            insert_commands.append(insert_sql_part01 + '\'' + val_str + '\',' + OPRT_DTTM + ');')

            # insert values for PR
            val_str = "', '".join([TRD_DT, TTL_BRCH_CD, TTL_BRCH_NM, TOT_EMP_CNT, TTL_OPER_CD_P, SE_TRD_FEE_P, DE_TRD_VAL_P, OPRT_IP_ADR])
            if "=" in SE_TRD_FEE_P:
                SE_TRD_FEE_P_REAL = SE_TRD_FEE_P[1:]
                val_str = "', '".join([TRD_DT, TTL_BRCH_CD, TTL_BRCH_NM, TOT_EMP_CNT, TTL_OPER_CD_P])
                val_str = val_str + "', " + SE_TRD_FEE_P_REAL + ", " + DE_TRD_VAL_P + ", '" + OPRT_IP_ADR 
            insert_commands.append(insert_sql_part01 + '\'' + val_str + '\',' + OPRT_DTTM + ');')  

        if "WS" in sheet["C66"].value:
            TTL_BRCH_CD = 'WHOLESALE'
            TTL_BRCH_NM = 'WS'
            bracket_pos_l = str(sheet["C66"].value).find('(')
            bracket_pos_r = str(sheet["C66"].value).find(')')
            TOT_EMP_CNT = (str(sheet["C66"].value))[bracket_pos_l+1:bracket_pos_r]
            # type
            TTL_OPER_CD_B = str(sheet["E66"].value).strip()
            TTL_OPER_CD_P = str(sheet["E67"].value).strip()
            # se daily value
            SE_TRD_FEE_B = str(sheet["F66"].value)
            SE_TRD_FEE_P = str(sheet["F67"].value)

            #de daily value
            DE_TRD_VAL_B = str(sheet["J66"].value)
            DE_TRD_VAL_P = str(sheet["J67"].value)
            if DE_TRD_VAL_B == 'None':
                DE_TRD_VAL_B = '0'
            if DE_TRD_VAL_P == 'None':
                DE_TRD_VAL_P = '0'

            # insert values for Broker
            val_str = "', '".join([TRD_DT, TTL_BRCH_CD, TTL_BRCH_NM, TOT_EMP_CNT, TTL_OPER_CD_B, SE_TRD_FEE_B, DE_TRD_VAL_B, OPRT_IP_ADR])
            if "=" in SE_TRD_FEE_B:
                            SE_TRD_FEE_B_REAL = SE_TRD_FEE_B[1:]
                            val_str = "', '".join([TRD_DT, TTL_BRCH_CD, TTL_BRCH_NM, TOT_EMP_CNT, TTL_OPER_CD_B])
                            val_str = val_str + "', " + SE_TRD_FEE_B_REAL + ", " + DE_TRD_VAL_B + ", '" + OPRT_IP_ADR
            insert_commands.append(insert_sql_part01 + '\'' + val_str + '\',' + OPRT_DTTM + ');')

            # insert values for PR
            val_str = "', '".join([TRD_DT, TTL_BRCH_CD, TTL_BRCH_NM, TOT_EMP_CNT, TTL_OPER_CD_P, SE_TRD_FEE_P, DE_TRD_VAL_P, OPRT_IP_ADR])
            if "=" in SE_TRD_FEE_P:
                SE_TRD_FEE_P_REAL = SE_TRD_FEE_P[1:]
                val_str = "', '".join([TRD_DT, TTL_BRCH_CD, TTL_BRCH_NM, TOT_EMP_CNT, TTL_OPER_CD_P])
                val_str = val_str + "', " + SE_TRD_FEE_P_REAL + ", " + DE_TRD_VAL_P + ", '" + OPRT_IP_ADR 
            insert_commands.append(insert_sql_part01 + '\'' + val_str + '\',' + OPRT_DTTM + ');') 

        logging.info("delete RP01N002 where TRD_DT =  \'{0}\'".format(TRD_DT))
        for cmd in insert_commands:
            logging.info(cmd)
            #print (cmd)
        