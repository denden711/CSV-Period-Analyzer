import pandas as pd
import numpy as np
from scipy.fft import fft, fftfreq
import os
from tkinter import Tk, filedialog, simpledialog
from openpyxl import Workbook
import logging

# ログ設定
logging.basicConfig(filename='period_calculation.log', level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')

def calculate_period(file_path, x_col_num=3, y_col_num=4, encoding='shift_jis'):
    try:
        # 指定されたエンコーディングでCSVファイルを読み込む
        try:
            df = pd.read_csv(file_path, encoding=encoding)
        except UnicodeDecodeError:
            logging.warning(f"{file_path} の {encoding} エンコーディングでエラーが発生しました。utf-8 を試みます。")
            df = pd.read_csv(file_path, encoding='utf-8')
        except Exception as e:
            logging.error(f"{file_path} の {encoding} エンコーディングで読み込みに失敗しました: {e}")
            return None

        # '#DIV/0!' や NaN を適切な NaN 値に置き換え、不要な列を削除
        df.replace('#DIV/0!', pd.NA, inplace=True)
        df.dropna(axis=1, how='all', inplace=True)
        df.columns = [col.strip() for col in df.columns]
        
        # 指定されたインデックスで列を抽出
        try:
            x_column = df.columns[x_col_num]
            y_column = df.columns[y_col_num]
        except IndexError as e:
            logging.error(f"{file_path} の列インデックスエラー: {e}")
            return None
        
        # x または y の値が NaN である行を削除
        df = df.dropna(subset=[x_column, y_column])

        # FFT解析のために y 値を抽出
        y_values = df[y_column].values

        # サンプルポイントの数
        N = len(y_values)

        # サンプル間隔
        T = df[x_column].diff().dropna().mean()

        if N < 2 or T == 0:
            raise ValueError("データポイントが不足しているか、サンプル間隔が無効です。")

        # FFTを実行
        yf = fft(y_values)
        xf = fftfreq(N, T)[:N//2]

        # ピーク周波数を特定
        peak_freq = xf[np.argmax(np.abs(yf[:N//2]))]

        # 周期を計算
        period = 1 / peak_freq
        
        return period
    except ValueError as e:
        logging.error(f"{file_path} の値エラー: {e}")
        return None
    except Exception as e:
        logging.error(f"{file_path} の処理中に予期しないエラーが発生しました: {e}")
        return None

def main():
    # Tkinter ルートウィンドウを作成
    root = Tk()
    root.withdraw()  # ルートウィンドウを隠す

    # ユーザーにディレクトリを選択させる
    directory = filedialog.askdirectory(title="CSVファイルが含まれるディレクトリを選択")

    if not directory:
        print("ディレクトリが選択されていません。")
        return

    # エンコーディングをユーザーに入力させる
    encoding = simpledialog.askstring("入力", "ファイルエンコーディングを入力（デフォルト: shift_jis）:", initialvalue="shift_jis")

    # Excelワークブックを準備
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(["ファイル名", "周期"])

    # ディレクトリ内の各CSVファイルを処理
    for file_name in os.listdir(directory):
        if file_name.endswith('.csv'):
            file_path = os.path.join(directory, file_name)
            period = calculate_period(file_path, encoding=encoding)
            if period is not None:
                sheet.append([file_name, period])
            else:
                sheet.append([file_name, "エラー"])
                logging.error(f"{file_name} の処理中にエラーが発生しました。")

    # 結果をExcelファイルに保存
    output_file = os.path.join(directory, "periods.xlsx")
    workbook.save(output_file)
    logging.info(f"結果が {output_file} に保存されました。")
    print(f"結果が {output_file} に保存されました。")

if __name__ == "__main__":
    main()
