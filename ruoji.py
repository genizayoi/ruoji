import pandas as pd
import openpyxl
import sys
import os
class StrCutter:
    def __init__(self, str_):
        self.str_ = str_
    def strCut(self, target_str_):
        if self.str_.find(target_str_) != -1:
            str_return_ = self.str_[:self.str_.find(target_str_)]
            self.str_ =self.str_[self.str_.find(target_str_) + 1:]
        else:
            str_return_=""
            self.str_=""
        return str_return_
class FileReader:
    def __init__(self, type_):
        self.type_ = type_
    def dfReturn(self, file_name_):
        if self.type_ == "csv":
            return pd.read_csv(path_+"/"+file_name_,header=None,skip_blank_lines = False,encoding=format_type_)
        elif self.type_ == "xlsx":
            return pd.read_excel(path_+"/"+file_name_,header=None,skip_blank_lines = False,encoding=format_type_)
path_ = os.getcwd()
files_ = os.listdir(path_)
print("Path:",path_)
print("Files:",files_)
while True:
    print("------------------------------------------------------------")
    print("BEFORE PROCEEDING, PLEASE CLOSE ALL DATA FILES")
    operation_code_ = StrCutter(input('Enter operation code: '))
    if operation_code_.str_.find("!") != -1:
        format_ = operation_code_.strCut("!")
        if format_ != '':
            print("Format:",format_)
            table_xy_ = StrCutter(operation_code_.strCut("!"))
            print("Table coord:",table_xy_.str_)
            data_xy_ = StrCutter(operation_code_.strCut("!"))
            print("Data coord:",data_xy_.str_)
            direction_ = operation_code_.strCut("!")
            if direction_!= "v":
                direction_="h"
            print("Direction:",direction_)
            format_type_ = operation_code_.strCut("!")
            if format_type_== "":
                format_type_="utf-8"
            print("Format type:",format_type_)
            table_xy_list_=[]
            data_xy_list_=[]
            table_number=0
            data_number=0
            input_broken_=False
            while table_xy_.str_ !='':
                table_xy_list_.append([table_xy_.strCut("."),table_xy_.strCut("."),table_xy_.strCut("."),table_xy_.strCut(".")])
            while data_xy_.str_ !='':
                data_xy_list_.append([data_xy_.strCut("."),data_xy_.strCut("."),data_xy_.strCut("."),data_xy_.strCut(".")])
            if len(table_xy_list_)==0 or len(data_xy_list_)==0:
                print("Error:Table or Data can not be NULL.")
                continue
            for i in table_xy_list_:
                if i[0]=="" or i[1]=="" or i[2]=="" or i[3]=="":
                    print("Error:Table input error.")
                    input_broken_=True
                    break
                table_number+=(1+int(i[2]))
            for i in data_xy_list_:
                if i[0]=="" or i[1]=="" or i[2]=="" or i[3]=="":
                    print("Error:Data input error.")
                    input_broken_=True
                    break
                data_number+=(1+int(i[2]))
            if input_broken_ is True:
                continue
            if table_number != data_number:
                print("Error:The number between Data and Table does not match.")
                continue
            is_fist_time_=True
            range_broken_=False
            target_files_name_list_=[]
            table_list_=[]
            data_list_=[]
            data_list_v_=[]
            if format_ == "csv":
                f_r_=FileReader("csv")
            elif format_ == "xlsx":
                f_r_=FileReader("xlsx")
            else:
                print("Error:Do not support for this format.")
                continue
            for file_ in files_:
                if file_[-len(format_)-1:] == ("."+format_) and file_!="Result_.csv" and file_!="~$Result_.csv":
                    data_ = []
                    target_files_name_list_.append(file_[:-len(format_)-1])
                    df = f_r_.dfReturn(file_)
                    if is_fist_time_ is True:
                        is_fist_time_ = False
                        for xy_ in table_xy_list_:
                            if xy_[3] == "h":
                                k=[0,1]
                            else:
                                k=[1,0]
                            for i in range(0,int(xy_[2])+1):
                                if (int(xy_[0])+(i*k[0]))>len(df) or (int(xy_[1])+(i*k[1]))>len(df.columns):
                                    print("Error:Out of range(Table).")
                                    range_broken_=True
                                    break
                                else:
                                    table_list_.append(df.iloc[int(xy_[0])+(i*k[0])-1,int(xy_[1])+(i*k[1])-1])
                    for xy_ in data_xy_list_:
                        if xy_[3] == "h":
                            k=[0,1]
                        else:
                            k=[1,0]
                        for i in range(0,int(xy_[2])+1):
                            if (int(xy_[0])+(i*k[0]))>len(df) or (int(xy_[1])+(i*k[1]))>len(df.columns):
                                print("Error:Out of range(Data).")
                                range_broken_=True
                                break
                            else:
                                data_.append(df.iloc[int(xy_[0])+(i*k[0])-1,int(xy_[1])+(i*k[1])-1])
                    if range_broken_ is True:
                        break
                    data_list_.append(data_)
            if range_broken_ is True:
                continue
            if direction_ == "v":
                for i in range(0,len(table_list_)):
                    data_v_ = []
                    for j in data_list_:
                        data_v_.append(j[i])
                    data_list_v_.append(data_v_)
                result_=pd.DataFrame(data_list_v_,index=table_list_,columns=target_files_name_list_)
            else:
                result_=pd.DataFrame(data_list_,index=target_files_name_list_,columns=table_list_)
            result_.to_csv("Result_.csv",encoding="utf_8_sig")
            print("------------------------------------------------------------")
            print("Complete")
            break
        else:
            s=[71,65,79,32,89,79,85,32,87,69,73,32,65,73,32,90,72,65,79,32,82,79,78,71,32,82,79,78,71,32,62,51,60]
            for i in s:
                print(chr(i),end='')
    else:
        print("RUOJI v1.0 By Kasico 2019/10/28")
#example:
#csv!26.1.180.v.!26.2.180.v.!v!shift-jis!
#csv!26.1.180.v.!26.2.180.v.!h!shift-jis!
#xlsx!10.2.0.v.!10.3.0.v.!v!utf8!
#xlsx!10.2.0.v.!10.3.0.v.!h!utf8!
#xlsx!14.9.1.h.!10.3.0.v.9.3.0.h.!v!utf8!