#!/usr/bin/env python
# coding: utf-8

# borrowed from https://qiita.com/chanmaru/items/8e5ebf7d8b0b21c8fd3a and  https://qiita.com/chanmaru/items/1b64aa91dcd45ad91540

# 入力directory(配下のCSVから複数選択）⇒出力directory とする。ディレクトリを再帰的にはたどらない
#
# 入力ファイル     >> [.........]  [Ref]
# 出力ディレクトリ >> [.........]  [Ref]
# [Start] [Cancel]
# [message area ]
#
# 入力ファイル     ⇒ pathlib.Path() を経由して targets 変数に入れる
# 出力ディレクトリ ⇒ pathlib.Path() を経由して outdir 変数に入れる

# TODO input files のフレームのサイズに上限を付けたい。というかスクロールバーつけたい。何も操作できなくなる。
# TODO pythonでwin binary化してコミット
# TODO massage フレーム付けたい。今どこまでやっているのか知りたい。もっとも、とても速いので気にならないかも

import os,sys
from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox
import pathlib
import csv2eml
# 参照ボタンのイベント
# infiles_REFERENCE_buttonクリック時の処理
def infiles_REFERENCE_button_clicked():

    #fTyp = [("","*")]
    # CSVを選択
    fTyp = [("","*.csv")]

    iDir = os.path.abspath(os.path.dirname(__file__))
    #filepath = filedialog.askopenfilename(filetypes = fTyp,initialdir = iDir)
    #file1.set(filepath)

    # ファイルリストを選択
    infiles = filedialog.askopenfilenames(filetypes = fTyp,initialdir = iDir)
    infiles_text = "\n".join(infiles) # \n 区切りでくっつける
    infiles_entry_text.set(infiles_text)

    # ディレクトリを選択
    # dir = tkinter.filedialog.askdirectory(initialdir = iDir)


# outdir_REFERENCE_buttonクリック時の処理
def outdir_REFERENCE_button_clicked():

    fTyp = [("","*")]
    # CSVを選択
    # fTyp = [("","*.csv")]

    iDir = os.path.abspath(os.path.dirname(__file__))

    #filepath = filedialog.askopenfilename(filetypes = fTyp,initialdir = iDir)
    #file1.set(filepath)

    # ファイルリストを選択
    # file = tkinter.filedialog.askopenfilenames(filetypes = fTyp,initialdir = iDir)
    # list = list(file)

    # ディレクトリを選択
    outdir = filedialog.askdirectory(initialdir = iDir)
    outdir_entry_text.set(outdir)


# buttonSTARTクリック時の処理
def buttonSTART_clicked():
    #infiles_label_textとoutdir_label_text を取り出して、
    #それぞれ pathlib オブジェクトを作ってcsv2eml
    infiles = infiles_entry_text.get().split("\n")
    outdir = pathlib.Path( outdir_entry_text.get() )
    for f in infiles:
        csv2eml.csv2eml(pathlib.Path(f), outdir)



if __name__ == '__main__':
    ###
    # rootの作成
    ###
    root = Tk()
    root.title('CSV2EML')
    root.resizable(False, False)

    ###
    ### infiles_frameの作成
    ###
    infiles_frame = ttk.Frame(root, padding=10)
    infiles_frame.grid()

    # 参照ボタンの作成
    infiles_REFERENCE_button = ttk.Button(root, text=u'Reference', command=infiles_REFERENCE_button_clicked)
    infiles_REFERENCE_button.grid(row=0, column=3)

    # ラベルの作成
    # 「ファイル」ラベルの作成
    infiles_label_text = StringVar()
    infiles_label_text.set('Input Files>>')
    infiles_label = ttk.Label(infiles_frame, textvariable=infiles_label_text)
    infiles_label.grid(row=0, column=0)

    # 参照ファイルパス表示ラベルの作成
    infiles_entry_text = StringVar()
#    infiles_entry = ttk.Entry(infiles_frame, textvariable=infiles_entry_text, width=50)
    infiles_entry = ttk.Label(infiles_frame, textvariable=infiles_entry_text, width=50)
    infiles_entry.grid(row=0, column=2)

    ###
    ### outdir_frameの作成
    ###
    outdir_frame = ttk.Frame(root, padding=10)
    outdir_frame.grid()

    # 参照ボタンの作成
    outdir_REFERENCE_button = ttk.Button(root, text=u'Reference', command=outdir_REFERENCE_button_clicked)
    outdir_REFERENCE_button.grid(row=1, column=3)

    # ラベルの作成
    # 「ファイル」ラベルの作成
    outdir_label_text = StringVar()
    outdir_label_text.set('Output dir>>')
    outdir_label = ttk.Label(outdir_frame, textvariable=outdir_label_text)
    outdir_label.grid(row=1, column=0)

    # 参照ファイルパス表示ラベルの作成
    outdir_entry_text = StringVar()
#    outdir_entry = ttk.Entry(outdir_frame, textvariable=outdir_entry_text, width=50)
    outdir_entry = ttk.Label(outdir_frame, textvariable=outdir_entry_text, width=50)
    outdir_entry.grid(row=1, column=2)


    ###
    ### control_frameの作成
    ###
    control_frame = ttk.Frame(root, padding=(0,5))
    control_frame.grid(row=2)

    # Startボタンの作成
    buttonSTART = ttk.Button(control_frame, text='Start', command=buttonSTART_clicked)
    buttonSTART.pack(side=LEFT)

    # Cancelボタンの作成
    buttonCANCEL = ttk.Button(control_frame, text='Cancel', command=quit)
    buttonCANCEL.pack(side=LEFT)

    ###
    ### 起動
    ###
    root.mainloop()
