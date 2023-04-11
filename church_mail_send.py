import win32com.client as win32
import pandas as pd
import os
import fnmatch
import shutil

hwp_path = os.getcwd() + '\성광교회 메일출력 파이썬.hwp'
excel_path = os.getcwd() + '\새신자관리표.xlsx'
df = pd.read_excel(excel_path)
now_path = os.getcwd()
os.mkdir(os.getcwd() + "\\자료")
save_path = os.getcwd() + "\\자료"

df = df.query("비고 != '완료'")
df = df.query("비고 != '반송'")

# 아래아한글 파일 열기
hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
hwp.Open(hwp_path)
hwp.XHwpWindows.Item(0).Visible = True


file_list = []
# 데이터 삽입
for i in range(len(df)):
    hwp.MoveToField("이름")
    hwp.PutFieldText("이름", df.iloc[i]["이름"])
    hwp.MoveToField("주소")
    hwp.PutFieldText("주소", df.iloc[i]["주소"])
    hwp.MoveToField("우편번호1")
    hwp.PutFieldText("우편번호1", df.iloc[i]["우편번호1"])
    hwp.MoveToField("우편번호2")
    hwp.PutFieldText("우편번호2", df.iloc[i]["우편번호2"])
    hwp.MoveToField("우편번호3")
    hwp.PutFieldText("우편번호3", df.iloc[i]["우편번호3"])
    hwp.MoveToField("우편번호4")
    hwp.PutFieldText("우편번호4", df.iloc[i]["우편번호4"])
    hwp.MoveToField("우편번호5")
    hwp.PutFieldText("우편번호5", df.iloc[i]["우편번호5"])
    hwp.MoveToField(None)
    new_filename = save_path +"\\"+df.iloc[i]["이름"]+".hwp"
    hwp.SaveAs(new_filename)

# 아래아한글 파일 저장 및 닫기

hwp.Quit()

hwp2 = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")

dir1=save_path
filename=fnmatch.filter(os.listdir(dir1),'*.hwp')

# os.chdir(save_path)
i=0
for file in filename:
    if i==0:
        hwp2.Open(save_path+"\\"+file, "HWP", "forceopen:true")
        hwp2.MovePos(3, 0, 0)
    else:
        hwp2.HAction.GetDefault("InsertFile", hwp2.HParameterSet.HInsertFile.HSet);
        option=hwp2.HParameterSet.HInsertFile
        option.filename = save_path+"\\"+file
        option.KeepSection = 1;
        option.KeepCharshape = 1;
        option.KeepParashape = 1;
        option.KeepStyle = 1;
        hwp2.HAction.Execute("InsertFile", hwp2.HParameterSet.HInsertFile.HSet);
        hwp2.MovePos(3, 0, 0)
    i=i+1


hwp2.HAction.GetDefault("FileSaveAs_S", hwp2.HParameterSet.HFileOpenSave.HSet);
option=hwp2.HParameterSet.HFileOpenSave
option.Attributes = 0;
option.filename = save_path+"메일머지.hwp"
option.Format = "HWP";
hwp2.HAction.Execute("FileSaveAs_S", hwp2.HParameterSet.HFileOpenSave.HSet);

hwp2.Quit()

shutil.rmtree(save_path)