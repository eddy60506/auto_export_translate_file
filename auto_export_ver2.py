# import module
import tkinter as tk
import os,time
import xlwings as xw
import threading
import win32gui,win32con,win32com

# setup function
def Upload_Finish():    # 上傳成功提示框
    WIN_FINISH = tk.Tk()
    WIN_FINISH.title("Success box")
    WIN_FINISH.geometry("200x100+1200+100")
    WIN_FINISH.config(bg="skyblue")
    WIN_FINISH.attributes("-topmost", 1)
    Generate_Basic_Label(WIN_FINISH,TEXT="上傳成功!!!", FG="yellow", BG="black", FONT="微軟正黑體 20")
    WIN_FINISH.mainloop()
# 生成按鈕
def Generate_Two_Btn(WIN,BTN1_TEXT="",BTN2_TEXT="",FONT="微軟正黑體 15",ABG="red",W="20",H="3",BTN1_CMD="",BTN2_CMD="",X1="0",Y1="0",X2="0",Y2="0"):
    BTN_ONE = tk.Button(WIN,text=BTN1_TEXT, font=FONT, activebackground=ABG, width=W, height=H, command =BTN1_CMD)
    BTN_ONE.place(x=X1,y=Y1)
    BTN_TWO = tk.Button(WIN,text=BTN2_TEXT, font=FONT, activebackground=ABG, width=W, height=H, command = BTN2_CMD)
    BTN_TWO.place(x=X2,y=Y2)
# 生成文字標籤
def Generate_Basic_Label(WIN,TEXT="", FG="black", BG="black", FONT="微軟正黑體 15",X="0",Y="0"): 
    LAB_BLANK = tk.Label(WIN,text=TEXT, fg=FG, bg=BG, font=FONT)
    # LAB_BLANK.grid()
    LAB_BLANK.place(x=X,y=Y)

def Press_All(): # 點擊全部上傳
    print("upload all files")
    print(XLSM_LIST)
    # 生成確認視窗
    WIN3 = tk.Tk()
    WIN3.title("Confirm box")
    WIN3.geometry("450x800+900+70")
    WIN3.config(bg="skyblue")
    WIN3.attributes("-topmost", 1) 
    # 確認視窗畫面設定
    Generate_Basic_Label(WIN3,TEXT="是否確定要上傳全部檔案?",FG="red",BG="skyblue", FONT="微軟正黑體 15",X=100)
    Generate_Basic_Label(WIN3,TEXT="="*30,FG="black",BG="skyblue", FONT="微軟正黑體 15",Y=25)

    XLSM_AMOUNT = 0     
    X = 0               # XLSM_LIST X 軸初始值設定
    Y = 50              # XLSM_LIST Y 軸初始值設定
    for x in XLSM_LIST: # 顯示所有 xlsm 
        if XLSM_AMOUNT > 16:
            Generate_Basic_Label(WIN3,TEXT= str(XLSM_AMOUNT+1) + ". " + x,FG="black",BG="white", FONT="微軟正黑體 9",X=X + 220,Y=Y - 17*25)
        else:
            Generate_Basic_Label(WIN3,TEXT= str(XLSM_AMOUNT+1) + ". " + x,FG="black",BG="white", FONT="微軟正黑體 9",X=X,Y=Y)
        XLSM_AMOUNT = XLSM_AMOUNT + 1
        Y = Y + 25

    Generate_Basic_Label(WIN3,TEXT="="*30,FG="black",BG="skyblue", FONT="微軟正黑體 15",Y=19*25)
    Generate_Basic_Label(WIN3,TEXT="總共選取 " + str(XLSM_AMOUNT) + " 個檔案",FG="white",BG="black", FONT="微軟正黑體 15",X=130,Y=20*25)
    Generate_Two_Btn(WIN3,BTN1_TEXT="全部上傳",BTN2_TEXT="取消",BTN1_CMD=lambda : Upload_Translate(XLSM_LIST),BTN2_CMD=lambda : Cancel(WIN3),X1=100,Y1=21*25+5,X2=100,Y2=25*25)
    WIN3 = tk.mainloop()

def Press_Check():  # 點擊部分上傳
    if len(CHECK_LIST) > 0: # 如果有選檔案
        print("Renew Check List:")
        print(CHECK_LIST)
        # 生成部分上傳確認畫面
        WIN2 = tk.Tk()
        WIN2.title("Confirm box")
        WIN2.geometry("450x800+900+70")
        WIN2.config(bg="skyblue")
        WIN2.attributes("-topmost", 1) 
        # 部分上傳確認畫面設定
        Generate_Basic_Label(WIN2,TEXT="是否確定要上傳這些檔案?\n",FG="red",BG="skyblue", FONT="微軟正黑體 15",X=100)
        
        CHECK_AMOUNT = 0 
        X = 0           # CHECK_LIST X 軸初始值設定
        Y = 50          # CHECK_LIST Y 軸初始值設定
        for x in CHECK_LIST:
            if CHECK_AMOUNT > 16:
                Generate_Basic_Label(WIN2, TEXT= x, FG="black", BG="white", FONT="微軟正黑體 9", X = X + 220, Y = Y - 17*25)
            else:
                Generate_Basic_Label(WIN2, TEXT= x , FG="black", BG="white", FONT="微軟正黑體 9", X = X, Y = Y)
            CHECK_AMOUNT = CHECK_AMOUNT + 1
            Y = Y + 25

        Generate_Basic_Label(WIN2,TEXT="="*50,FG="black",BG="skyblue", FONT="微軟正黑體 9",Y=19*25)
        Generate_Basic_Label(WIN2,TEXT="總共選取 " + str(CHECK_AMOUNT) + " 個檔案",FG="white",BG="black", FONT="微軟正黑體 15",X=130,Y=20*25)
        Generate_Two_Btn(WIN2,BTN1_TEXT="確定上傳",BTN2_TEXT="取消",BTN1_CMD=lambda : Upload_Translate(CHECK_LIST),BTN2_CMD=lambda : Cancel(WIN2),X1=100,Y1=21*25+5,X2=100,Y2=25*25)
        Generate_Basic_Label(WIN2,TEXT="注意：點擊上傳後請等待\n「上傳成功」視窗跳出再關閉工具",FG="red",BG="black", FONT="微軟正黑體 15",X=70,Y=29*25)
        WIN2.mainloop()

    else:   # no file checked
        WIN4 = tk.Tk()
        WIN4.title("ERROR!!!")
        WIN4.geometry("450x50+1200+400")
        WIN4.config(bg="black")
        WIN4.attributes("-topmost", 1)

        Generate_Basic_Label(WIN4,TEXT="You didn't chose any file!!!Please check agian!!!",FG="red",BG="skyblue", FONT="微軟正黑體 15")

        WIN4.mainloop()

def Cancel(BTN):    # 點擊取消 
    print("Cancel!!")
    BTN.destroy()

def Readstatus_value(KEY):      # 判斷 checkbutton state
    var_obj = VAR.get(KEY)
    if var_obj.get() == 1:      # var_obj.get() = 1 有勾選
        CHECK_LIST.append(KEY)  # 加入list
    else :                      # var_obj.get() = 0 沒勾選
        CHECK_LIST.remove(KEY)  # 從list中刪除

def find_process(name): # find_process("Exporter.exe")
    # 取得 Excel COMObject
    objWMIService = win32com.client.Dispatch("WbemScripting.SWbemLocator")
    objSWbemServices = objWMIService.ConnectServer(".", "root\cimv2")
    colItems = objSWbemServices.ExecQuery(
         "Select * from Win32_Process where Caption = '{0}'".format(name))
    return len(colItems)    # 若目標程序不存在 return = 0

def close_win():
    while True:
        time.sleep(3)
        w = win32gui.FindWindow(None, 'Microsoft Excel')
        win32gui.PostMessage(w, win32con.WM_CLOSE,0,0)
        e = win32gui.FindWindow(None, 'Exporter')
        win32gui.PostMessage(e, win32con.WM_CLOSE,0,0)
        #print('close windows')
    return

def Upload_Translate(file):
    if __name__ == '__main__':
        
        if len(file) == 0:
            print("no file to upload!!")    # 取得 當下目路內所有檔名 (list)
        else:
            UPLOAD_FILE_LIST = file

        app = xw.App(visible=True)  # 取得 Excel worksheet

        # close windows
        r = threading.Thread(target=close_win)
        r.daemon = True
        r.start()
        
        print("\n==== Translate Upload Start ====")
        
        for x in UPLOAD_FILE_LIST:
            if x[-4:] == 'xlsm':    # 若副檔名 == 'xlsm', [-4:] 從倒數第 4 個字元到最後一個
                wb = xw.Book(x)
                wb.macro('Sheet1.CommandButton1_Click')()
                print(x)
                time.sleep(1)
                while find_process("Exporter.exe"): # 如果發現 exporter 還在運作 ， 等 1 秒，直到 exporter 結束
                                                    # 上傳完畢就關閉 workbook
                    time.sleep(1)
                wb.close()
                #print(x, 'close')
                time.sleep(2)
        print("==== Translate Upload finished ====")
        app.kill()
        Upload_Finish()
        # input("Press ENTER to exit")

# start main progress

# tkinter gui setup

WIN = tk.Tk()                           # tkinter 視窗物件
WIN.title("Auto Upload Translate")      # title 設定
WIN.geometry("600x800+1200+100")        # 視窗位置設定 geometry(長x寬+左邊界距+上邊界距)
WIN.config(bg="black")                  # bg=background 顏色設定
WIN.attributes("-topmost", 0)           # 視窗置頂設置 attributes("-topmost", 0) 0 = False, 1 = True
# WIN.attributes("-toolwindow", 1) 

# variable setup
VAR = dict()
CHECK_LIST = []
XLSM_LIST = []
CHECKBOX_GEO = "300x870+900+70"

FILE_LIST = os.listdir()

# 主畫面設定

Generate_Basic_Label(WIN,TEXT="請勾選要上傳的翻譯檔\n或是直接點擊「全部上傳」",FG="skyblue",BG="black", FONT="微軟正黑體 15",X=170,Y=0)
Generate_Basic_Label(WIN,TEXT="="*40,FG="red",BG="black",X=0,Y=50)

for x in FILE_LIST:         # 排除非 xlsm 檔案
    if x[-4:] == 'xlsm':
        XLSM_LIST.append(x)

FILE_AMOUNT = 0
X = 1       # XLSM_LIST X 軸初始值設定
Y = 80      # XLSM_LIST Y 軸初始值設定
for CHILD in XLSM_LIST:     # 動態生成 checkbutton
    VAR[CHILD]=tk.IntVar()  # IntVar() 用來判斷 checkbutton state
    CHK = tk.Checkbutton(WIN,text=str(FILE_AMOUNT+1)+ ". " + CHILD, variable=VAR[CHILD], command=lambda KEY=CHILD: Readstatus_value(KEY))
    # 利用lamba func 讓 command 傳遞引數
    if FILE_AMOUNT>16:  
        CHK.place(x=X+300, y=Y-17*25)
    else:  
        CHK.place(x=X, y=Y)
    FILE_AMOUNT=FILE_AMOUNT+1
    Y = Y + 25

Generate_Basic_Label(WIN,TEXT="="*40,FG="red",BG="black",X=0,Y=50+18*25)
Generate_Basic_Label(WIN,TEXT="總共有 " + str(FILE_AMOUNT) + " 個檔案",FG="white",BG="black", FONT="微軟正黑體 15",X=200,Y=50+19*25)
Generate_Two_Btn(WIN,BTN1_TEXT="上傳部份翻譯檔",BTN2_TEXT="全部上傳",BTN1_CMD=Press_Check,BTN2_CMD=Press_All,X1=160,Y1=5+50+20*25,X2=160,Y2=5+50+24*25)
# 修改聲明 & 資訊
Generate_Basic_Label(WIN,TEXT="\n\nModified by X-Legend Eddy 2020 Oct Email: eddy60506@x-legend.com.tw",FG="white",BG="black",FONT="微軟正黑體 9",X=70,Y=70+27*25)

WIN.mainloop()  # tkinter

# 參考資料:
# https://www.delftstack.com/zh-tw/howto/python-tkinter/how-to-pass-arguments-to-tkinter-button-command/
# https://stackoverflow.com/questions/24663661/tkinter-get-values-from-dynamic-checkboxes
# https://pypi.org/project/auto-py-to-exe/