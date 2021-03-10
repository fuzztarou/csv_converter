import os
import openpyxl
import tkinter as tk
import tkinter.filedialog as tkdialog
import tkinter.messagebox as tkmbox
import module_spfi_csv_cov_pd_rev01 as csvconv

root = tk.Tk()          #ウィンドウオブジェクトを作成
root.withdraw()         #ウィンドウを表示しない設定
fTyp = [("","csv")]                                                     #ファイルタイプ
iDir = os.path.abspath(os.path.dirname(__file__))                       #ファイルのdir名からdirパスを取得
tkmbox.showinfo('csv変換','変換するCSVファイルを選択してください(複数可)')

files = tkdialog.askopenfilenames(filetypes = fTyp, initialdir = iDir)  #ファイルのフルパスのタプルを作成 

for file in files:
    csvconv.spfi_csv_conv(file)