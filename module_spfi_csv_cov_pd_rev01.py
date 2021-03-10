import os
import openpyxl
import pandas as pd
import matplotlib.pyplot as plt


def spfi_csv_conv (csvfname):

    #csvファイルの読み込み
    df = pd.read_csv(csvfname, header = 0,                      #ヘッダなし
                    delimiter = ";",                            #デリミタ指定
                    dtype = float)                              #型指定

    #読み込んだcsvファイルのカラム名の調整
    df = df.add_suffix("(Hz)")                                  #カラム名に(Hz)を追加
    new_df = df.rename(columns={"0.000(Hz)":"On Time (ms)"})    #カラム１の名前を変更している
    #この後はnew_dfを使う                                         #renameすると新しいオブジェクトが返されるので

    i = 0                                                       #forループ用
    l_columns = list(new_df.columns)                            #カラム名のリスト作成

    #(カラム数−1)回ループを回す
    for i in range (len(new_df.columns) - 1):
        if i == 0:
            #axeオブジェクトをaxに代入している
            ax = new_df.plot(x=l_columns[0],                    #xカラム名で指定
                                y=l_columns[i+1],               #yをカラム名で指定
                                marker="o",                     #マーカーを指定
                                linestyle="-",                  #線を指定
                                figsize=(5,4))                  #グラフサイズを指定
        else:
            new_df.plot(x=l_columns[0],
                                y=l_columns[i+1],
                                marker="o",
                                linestyle="-",
                                ax=ax)                          #追加元のaxeオブジェクトを指定

    #拡張子なしのファイル名を取得 ([ファイル名, 拡張子]のリストからファイル名を取り出している)
    #splitextは拡張子とそれ以外のタプルを返す　[1]に拡張子
    filename = os.path.splitext(os.path.basename(csvfname))[0]

    #グラフの見た目の設定
    ax.set_title(filename)                               #グラフタイトル
    ax.set_xlabel("Turn On Time (ms)")                   #軸ラベルの設定
    ax.set_ylabel("Flow Rate (mg/stroke)")
    ax.grid(1)                                           #グリッドの設定      
    ax.set_xlim(0,16)                                    #軸の範囲の設定
    ax.set_ylim(-1,50)

    #公式　matplotlib.axe　https://matplotlib.org/3.3.4/api/axes_api.html
    #Qiita　早く知っておきたかったmatplotlibの基礎知識　https://qiita.com/skotaro/items/08dc0b8c5704c94eafb9

    plt.savefig("myplot.png")                           #グラフを一旦保存

    #ここからエクセルファイルを作成して書き込んでいく
    new_df.to_excel(filename + '.xlsx',                 #ファイル名
                    sheet_name=filename,                #シート名
                    index=False)                        #DataFrameのindexは書き込まない

    wb = openpyxl.load_workbook(filename + '.xlsx')     #エクセルファイルの読み込み
    ws = wb.active                                      #開いているシートをwsに代入
    ws.insert_rows(1, 22)                               #エクセルの1行目にx行挿入してグラフのスペースをあける
    img = openpyxl.drawing.image.Image('myplot.png')    #保存したグラフの読み込み
    ws.add_image(img, "A1")                             #エクセルに画像を挿入
    wb.save(filename + '.xlsx')                         #エクセルを保存する
