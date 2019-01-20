import ctypes
import array
import win32con
import win32gui
import subprocess
import sys
from time import sleep
from tkinter import filedialog
from tkinter import messagebox


AKANE_EXE=""#実行ファイルのpath
AKANE_TITLE="VOICEROID＋ 琴葉茜"#ウィンドウのタイトル

AOI_EXE=""
AOI_TITLE="VOICEROID＋ 琴葉葵"

NUM_VOICEROID=2
CHANGE_TIME=0.18  #再生ボタンが一時停止になるまで待機する時間,台詞が飛ぶ場合は大きくすること
START_UP_TIME=4 #起動待機時間(秒)
WAIT_TIME=0.02#台詞を言っている間に、言い終わっているか確認する間隔(秒)
SLEEP_TIME=0.5#改行のみの行があった時の待ち時間

def enum_child_windows_proc(handle, list):
    list.append(handle)
    return 1

def serchChildWin(Phandle):
    child_handles = array.array("i")
    ENUM_CHILD_WINDOWS = ctypes.WINFUNCTYPE(\
            ctypes.c_int,\
            ctypes.c_int,\
            ctypes.py_object)
    ctypes.windll.user32.EnumChildWindows(\
            Phandle,\
            ENUM_CHILD_WINDOWS(enum_child_windows_proc),\
            ctypes.py_object(child_handles))
    return child_handles

def initVoiceroid(path,title):#ボイロの実行ファイルのパスとウィンドウタイトル
    #ウィンドウタイトルex茜ちゃん)"VOICEROID＋ 琴葉茜"
    #本当は先に起動してるかチェックする必要あり?
    #でも先に起動してなんやかんやすると子ウィンドウ順番がずれてしまうので
    #起動してたらボイスロイドを再起動する必要あり
    subprocess.Popen(path)#起動！
    sleep(START_UP_TIME)#起動時間待機
    handle=win32gui.FindWindow(None,title)
    if handle==0:
        print("error")
        sys.exit()
    return handle


class Voiceroid:
    def __init__(self,hwnd):
        self.hwnd=hwnd
        child=serchChildWin(self.hwnd)
        self.SAISEI=child[11]#再生ボタンのウィンドウ
        self.TALK=child[9]#テキストをいれるウィンドウ
                               
        #引数4つ目、座標(200,15)を示している.上位16ビットにy,下位16ビットにx
        #x=15,y=200,y<16してx|yすると9832...になる
        win32gui.SendMessage(child[19],win32con.WM_LBUTTONDOWN,win32con.MK_LBUTTON,983240)
        win32gui.SendMessage(child[19],win32con.WM_LBUTTONUP,0,0)
    
        child=serchChildWin(child[19])#音声効果の子ウィンドウ取得
        self.YOKUYOU=child[24]#抑揚の値を変えるウィンドウ
        self.TAKASA=child[25]
        self.WASOKU=child[26]
        self.ONRYO=child[27]
        self.warikomi=0
    def setText(self,s):
        win32gui.SendMessage(self.TALK,win32con.WM_SETTEXT,0,s)

    def putVoice(self):#再生ボタンを押す
        win32gui.SendMessage(self.SAISEI,win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON,None)
        win32gui.SendMessage(self.SAISEI,win32con.WM_LBUTTONUP,None,None)
        #DOWN,UPでクリック。DOWNの第三引数はほかに同時に押すボタンがあれば|MK_..と書くらしい

    def waitVoice(self):#しゃべり終わるまで待機する
        sleep(CHANGE_TIME)#必要---------------------------------------環境によって待機時間が変わりそう？
        buf=ctypes.create_unicode_buffer(4)
        while True:
            win32gui.SendMessage(self.SAISEI,win32con.WM_GETTEXT,4,buf)
            if buf.value==" 再生":
                break
            sleep(WAIT_TIME)
        
    def setParam(self,volume,speed,height,intonation):
        if intonation>=0 and intonation<=2.0:
            win32gui.SendMessage(self.YOKUYOU,win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON,None)
            win32gui.SendMessage(self.YOKUYOU,win32con.WM_LBUTTONUP,None,None)
            win32gui.SendMessage(self.YOKUYOU,win32con.WM_SETTEXT,0,str(intonation))

        if height>=0.5 and height<=2.0:
            win32gui.SendMessage(self.TAKASA,win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON,None)
            win32gui.SendMessage(self.TAKASA,win32con.WM_LBUTTONUP,None,None)
            win32gui.SendMessage(self.TAKASA,win32con.WM_SETTEXT,0,str(height))

        if speed>=0.5 and speed<=4.0:
            win32gui.SendMessage(self.WASOKU,win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON,None)
            win32gui.SendMessage(self.WASOKU,win32con.WM_LBUTTONUP,None,None)
            win32gui.SendMessage(self.WASOKU,win32con.WM_SETTEXT,0,str(speed))

        if volume>=0.5 and volume<=2.0:
            win32gui.SendMessage(self.ONRYO,win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON,None)
            win32gui.SendMessage(self.ONRYO,win32con.WM_LBUTTONUP,None,None)
            win32gui.SendMessage(self.ONRYO,win32con.WM_SETTEXT,0,str(volume))
       #音声効果タブのパラメータを決定する関数、範囲外なら変更なし(常に初期値をもってるし大丈夫やろ
            
    def closeVoiceroid(self):
        win32gui.SendMessage(self.hwnd,win32con.WM_CLOSE,0,0)

    def talkVoice(self,s):
        self.setText(s)
        self.putVoice()
        self.waitVoice()
        
#--------------------ここまででボイロ操作できる------------------------------
#-------------------以下テキストを読ませるための準備------------------
def is_float(s):
    try:
        float(s)
    except:
        return False
    return True


def findEx(s,c,cc):
    #@ or ＠を探す。見つからないときは-1
    a=s.find(c)
    b=s.find(cc)
    if a==-1 and b==-1:
        return -1
    elif a==-1:
        a=b#先に見つかった方を処理するため
    elif a>b and b!=-1:
        a=b
    return a

def serchText(s):#findExと合わせて
    a=findEx(s,"@","＠")
    if a==-1:
        return s,"",""#抜き出すものがないよ
        
    b=findEx(s[a:],"：",":")+a
    if b==-1:
        return s,"",""

    #@以下,＠～：,：以降に三分割して出力
    if b==len(s):
        return s[:a],s[a:b+1],""
    else :
        return s[:a],s[a:b+1],s[b+1:]


def textCheck(s):#＠と：の位置と個数見るだけのプログラム簡易的なチェック
    #適当なのであまりあてにしない
    i=1
    error=0
    for ss in s:
        a=0
        b=0
        while a!=-1 and b!=-1:
            a=findEx(ss,"@","＠")
            b=findEx(ss,":","：")
            if a==-1 and b!=-1:
                print("error:"+str(i)+"行目")
                error+=1
                break
            elif a!=-1 and b==-1:
                print("error:"+str(i)+"行目")
                error+=1
                break
            elif a>b:
                print("error:"+str(i)+"行目")
                error+=1
                break
            else :
                if b==len(ss):
                    break
                else :
                    ss=ss[b+1:]
        i+=1
    print("-------")    
    return error   


def textPutout(s,voiro):#warikomi=2のときwait()しない
    if len(s)!=0:
        if voiro.warikomi==0:
            voiro.talkVoice(s)
            
        elif voiro.warikomi==1:
            voiro.waitVoice()
            voiro.talkVoice(s)
            voiro.warikomi=0
        else :
            voiro.setText(s)
            voiro.putVoice()
            voiro.warikomi=1
        print(s)
       


def textTranslation(s,voiro,num):#ボイロが切り替わる時に値を返す、それ以外はそのまま
    if s=="":
        return num
    if s.find("あかね")!=-1 or s.find("茜")!=-1 or s.find("akane")!=-1:
        print("<akanechan>")
        return 0

    elif s.find("あおい")!=-1 or s.find("葵")!=-1 or s.find("aoi")!=-1:
        print("<aoichan>")
        return 1
        
    elif s.find("ぱらめ")!=-1 or s.find("パラメ")!=-1:
        param=[-1.0,-1.0,-1.0,-1.0]
        a=findEx(s,"め","メ")
        for i in range(4):
            b=findEx(s[1:],",","、")+1
            if b==0:
                b=len(s)-1
            if is_float(s[a+1:b])==True:    
                param[i]=float(s[a+1:b])
            a=0
            s=s[b:]
        voiro.setParam(param[0],param[1],param[2],param[3])   
        print("<para="+str(param[0])+"|"+str(param[1])+"|"+str(param[2])+"|"+str(param[3])+">")
        

    elif s.find("ねる")!=-1:
        a=s.find("る")
        if is_float(s[a+1:len(s)-1])==True:
            print("<"+str(s[a+1:len(s)-1])+"秒休憩>")
            sleep(float(s[a+1:len(s)-1]))
        else :
            print("スリープエラー")

    elif s.find("わりこみ")!=-1:
        voiro.warikomi=3
        
    else :
        print("<＠～：の記述が正しくない>")
    return num    


def textRead(voice):#読み込みとチェック分けた方が良かった
    text=""
    s=""
    while text=="":
        typ=[('テキストファイル','*.txt')]
        dir="./"
        text=filedialog.askopenfilename(filetypes = typ, initialdir = dir) 
        if text=="":
            ret=messagebox.askyesno("角煮","テキストファイルが選択されてません\
終了しますか?")
            if ret==True:
                for i in range(NUM_VOICEROID):
                    voice[i].closeVoiceroid()
                sys.exit()
        else :
            f=open(text)
            s=f.readlines()
            f.close()
            error=textCheck(s)
            if error!=0:
                 ret=messagebox.askyesno("エラー","雑なチェックでは文が\
おかしいことがわかりました。無視して続けますか？")                            
                 if ret==False:
                    text=""
        return s            
#--------------------------------------------------------------------

selectVoice=0
actVoiceroid=[]

actVoiceroid.append(Voiceroid(initVoiceroid(AKANE_EXE,AKANE_TITLE)))
actVoiceroid.append(Voiceroid(initVoiceroid(AOI_EXE,AOI_TITLE)))


talking=True
while talking==True:
    print("-------------")
    s=textRead(actVoiceroid)
    for ss in s:
        result=[0,0,0]
        if ss[len(ss)-1]=="\n":
            ss=ss[:len(ss)-1]
        if ss=="":
            sleep(SLEEP_TIME)
            print("<改行のみの行がありました>")
        while result[2]!="":
            result=serchText(ss)
            textPutout(result[0],actVoiceroid[selectVoice])
            selectVoice=textTranslation(result[1],actVoiceroid[selectVoice],selectVoice)
            ss=result[2]
    talking=messagebox.askyesno("ｱｶﾈﾁｬﾝｶﾜｲｲﾔｯﾀｰ","他のテキストを読みま\
すか？いいえの場合は終了します")

for i in range(NUM_VOICEROID):
    actVoiceroid[i].closeVoiceroid()




        
