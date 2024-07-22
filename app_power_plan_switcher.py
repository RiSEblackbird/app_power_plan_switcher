# -*- coding: utf-8 -*-
# アプリ名: 0. 電源オプション
import subprocess
import sys
import traceback
import tkinter as tk
from tkinter import font
from tkinter import messagebox
import socket
import csv

# 定数としてウィンドウ位置情報を保存するファイル名を設定
POSITION_FILE = 'window_position_app_power_plan_switcher.csv'

# 定数としてホスト名を取得
HOSTNAME = socket.gethostname()

def get_exception_trace():
    '''例外のトレースバックを取得'''
    t, v, tb = sys.exc_info()
    trace = traceback.format_exception(t, v, tb)
    return trace

def save_position(root):
    """
    ウィンドウの位置のみをCSVファイルに保存する。異なるホストの情報も保持。
    """
    print("ウィンドウ位置を保存中...")
    # root.geometry()から位置情報のみを取り出す
    position_info = root.geometry().split('+')[1:]
    position_str = '+' + '+'.join(position_info)
    position_data = [HOSTNAME, position_str]
    existing_data = []
    try:
        # 既存のデータを読み込んで、現在のホスト以外の情報を保持
        with open(POSITION_FILE, newline='', encoding="utf_8_sig") as csvfile:
            reader = csv.reader(csvfile)
            existing_data = [row for row in reader if row[0] != HOSTNAME]
    except FileNotFoundError:
        print("ファイルが存在しないため、新規作成します。")
    
    # 現在のホストの情報を含む全データを書き込む
    existing_data.append(position_data)
    with open(POSITION_FILE, 'w', newline='', encoding="utf_8_sig") as csvfile:
        writer = csv.writer(csvfile)
        writer.writerows(existing_data)
    print("保存完了")

def restore_position(root):
    """
    CSVファイルからウィンドウの位置のみを復元する。
    """
    print("ウィンドウ位置を復元中...")
    try:
        with open(POSITION_FILE, newline='', encoding="utf_8_sig") as csvfile:
            reader = csv.reader(csvfile)
            for row in reader:
                if row[0] == HOSTNAME:
                    position_str = row[1].strip()  # 余計な空白がないか確認
                    if position_str.startswith('+'):
                        position_str = position_str[1:]  # 先頭の余分な '+' を取り除く
                    print(f"復元データ: {position_str}")
                    # サイズ情報なしで位置情報のみを設定
                    root.geometry('+' + position_str)
                    break
    except FileNotFoundError:
        print("位置情報ファイルが見つかりません。")

# 電源プランを変更する関数
def change_power_plan(plan_guid):
    try:
        result = subprocess.run(['powercfg', '/setactive', plan_guid], stdout=subprocess.PIPE, stderr=subprocess.PIPE, universal_newlines=True)
        if result.returncode == 0:
            print(f"電源プラン '{plan_guid}' が変更されました。")  # コンソールへの出力
        else:
            messagebox.showerror("エラー", f"パラメータが無効です: '{plan_guid}'\n{result.stderr}")
    except subprocess.CalledProcessError as e:
        messagebox.showerror("エラー", f"電源プランの変更に失敗しました: '{plan_guid}'\n{e}")


# GUIアプリケーションのクラス
class PowerPlanApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('電源オプション')
        self.geometry('250x350')
        self.listbox_guids = {}
        self.active_plan_guid = self.get_active_plan_guid().split()[0]  # ガイドのみを取得
        print("アクティブなプランのGUIDを取得: {}".format(self.active_plan_guid))
        self.create_widgets()

    def create_widgets(self):
        listbox_font = font.Font(family="Helvetica", size=14)
        self.listbox = tk.Listbox(self, font=listbox_font)
        self.listbox.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        self.populate_listbox()
        self.listbox.bind('<<ListboxSelect>>', self.on_select)

    def select_active_plan(self):
        # リストボックス内の全項目を確認し、アクティブなプランを選択する
        for index, name in enumerate(self.listbox.get(0, tk.END)):
            if self.listbox_guids.get(name) == self.active_plan_guid:
                self.listbox.selection_set(index)  # アクティブなプランを選択状態にする
                self.listbox.see(index)  # アクティブなプランが見えるようにスクロール
                break

    def get_active_plan_guid(self):
        # 現在アクティブな電源プランのGUIDを取得する
        result = subprocess.run(['powercfg', '/getactivescheme'], stdout=subprocess.PIPE, universal_newlines=True)
        active_guid = result.stdout.split(':')[1].strip()
        return active_guid

    def populate_listbox(self):
        result = subprocess.run(['powercfg', '/list'], stdout=subprocess.PIPE, universal_newlines=True)
        plan_list = []
        for line in result.stdout.split('\n'):
            if "電源設定の GUID:" in line:
                parts = line.split(' (')
                guid = parts[0].split(':')[1].strip()
                name = parts[1].rstrip(')').replace(')', '').replace('*', '').strip()  # 名前の整形
                plan_list.append((name, guid))
        
        # プラン名に基づいてリストをソート
        plan_list.sort()

        # ソートされたリストから項目をリストボックスに追加
        for name, guid in plan_list:
            self.listbox.insert(tk.END, name)
            self.listbox_guids[name] = guid
            print("プラン追加: 名前={}, GUID={}".format(name, guid))
            if guid == self.active_plan_guid:
                print("アクティブなプランをリストに追加: {}".format(name))
                self.listbox.selection_set(self.listbox.size() - 1)  # 最新の追加項目を選択

    def on_select(self, event):
        selection = event.widget.curselection()
        if selection:
            index = selection[0]
            name = event.widget.get(index)
            guid = self.listbox_guids.get(name)
            if guid:
                print("プラン変更: {}".format(name))
                change_power_plan(guid)

    def set_active_selection(self):
        for index, name in enumerate(self.listbox.get(0, tk.END)):
            if self.listbox_guids.get(name) == self.active_plan_guid:
                self.listbox.selection_set(index)
                self.listbox.see(index)
                print("アクティブなプランを選択: {}".format(name))  # 選択したアクティブなプランを表示
                break


# アプリケーションの終了時の処理をカスタマイズする
def on_close():
    save_position(app)  # ウィンドウの位置を保存
    app.destroy()  # ウィンドウを破壊する

# メインの処理
# メイン処理
try:
    app = PowerPlanApp()
    restore_position(app)
    app.protocol("WM_DELETE_WINDOW", on_close)  # 終了時処理の設定
    app.mainloop()

except Exception as e:
    t, v, tb = sys.exc_info()
    trace = traceback.format_exception(t, v, tb)
    print(trace)
