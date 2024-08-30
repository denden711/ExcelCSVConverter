import os
import pandas as pd
from tkinter import Tk, filedialog, messagebox
import logging

# ログ設定 (エンコーディングを指定)
logging.basicConfig(filename='excel_to_csv_converter.log', level=logging.INFO, 
                    format='%(asctime)s %(levelname)s: %(message)s', 
                    encoding='utf-8')

def select_directory(prompt="ディレクトリを選択してください"):
    """ユーザーにディレクトリを選択させるダイアログを表示"""
    Tk().withdraw()  # メインウィンドウを非表示にする
    directory = filedialog.askdirectory(title=prompt)
    return directory

def convert_excel_to_csv(input_directory, output_directory):
    """指定されたディレクトリ内のExcelファイルをCSVに変換する関数"""
    try:
        if not os.path.exists(input_directory):
            raise FileNotFoundError("指定された入力ディレクトリが存在しません。")
        
        if not os.path.exists(output_directory):
            raise FileNotFoundError("指定された出力ディレクトリが存在しません。")
        
        # 入力ディレクトリ内のすべてのファイルをチェック
        for filename in os.listdir(input_directory):
            if filename.endswith(".xlsx") or filename.endswith(".xls"):
                try:
                    # Excelファイルのパスを取得
                    excel_path = os.path.join(input_directory, filename)
                    # ファイル名と拡張子を分離
                    file_base_name = os.path.splitext(filename)[0]
                    # CSVファイルのパスを設定（出力ディレクトリに保存）
                    csv_path = os.path.join(output_directory, file_base_name + ".csv")

                    # Excelファイルを読み込む（シートが1枚だけと仮定）
                    try:
                        df = pd.read_excel(excel_path, engine='openpyxl')
                    except Exception as e:
                        logging.error(f"Excelファイル '{filename}' の読み込み中にエラーが発生しました: {e}")
                        messagebox.showerror("エラー", f"Excelファイル '{filename}' の読み込み中にエラーが発生しました。詳細はログを確認してください。")
                        continue

                    # CSVファイルに書き出す
                    try:
                        df.to_csv(csv_path, index=False, encoding='utf-8-sig')
                        logging.info(f"'{filename}' を '{file_base_name}.csv' として変換し、'{output_directory}' に保存しました。")
                    except Exception as e:
                        logging.error(f"CSVファイル '{file_base_name}.csv' の書き出し中にエラーが発生しました: {e}")
                        messagebox.showerror("エラー", f"CSVファイル '{file_base_name}.csv' の書き出し中にエラーが発生しました。詳細はログを確認してください。")
                        continue

                    messagebox.showinfo("成功", f"{filename} を {file_base_name}.csv に変換し、'{output_directory}' に保存しました。")

                except Exception as e:
                    logging.error(f"ファイル '{filename}' の変換中に予期しないエラーが発生しました: {e}")
                    messagebox.showerror("エラー", f"ファイル '{filename}' の変換中にエラーが発生しました。詳細はログを確認してください。")
    
    except FileNotFoundError as e:
        logging.error(f"ディレクトリエラー: {e}")
        messagebox.showerror("エラー", f"指定されたディレクトリが見つかりません。")
    
    except PermissionError as e:
        logging.error(f"アクセス権エラー: {e}")
        messagebox.showerror("エラー", f"ディレクトリまたはファイルへのアクセス権がありません。")
    
    except Exception as e:
        logging.error(f"ディレクトリ内のファイル処理中に予期しないエラーが発生しました: {e}")
        messagebox.showerror("エラー", f"ファイル処理中にエラーが発生しました。詳細はログを確認してください。")

def main():
    """プログラムのメイン関数"""
    # 入力ディレクトリ選択ダイアログを表示
    input_directory = select_directory("入力ディレクトリを選択してください")
    if not input_directory:
        messagebox.showwarning("警告", "入力ディレクトリが選択されていません")
        return
    
    # 出力ディレクトリ選択ダイアログを表示
    output_directory = select_directory("出力ディレクトリを選択してください")
    if not output_directory:
        messagebox.showwarning("警告", "出力ディレクトリが選択されていません")
        return
    
    # ExcelファイルをCSVに変換
    convert_excel_to_csv(input_directory, output_directory)
    messagebox.showinfo("完了", "すべてのファイルの変換が完了しました")

if __name__ == "__main__":
    main()
