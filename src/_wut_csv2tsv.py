# 推奨されるスクリプト名: convert_csv_to_tsv.py
# このスクリプトは、CSVファイルをドラッグ＆ドロップすることでTSV形式に変換します。
# CSVファイルをスクリプトにドロップすると、同じディレクトリに同名のTSVファイルが生成されます。

import pandas as pd
import sys
import os

def convert_csv_to_tsv(csv_file):
    """Convert CSV file to TSV format."""
    file_name, file_extension = os.path.splitext(csv_file)
    
    try:
        # エンコーディング自動判定用
        for encoding in ['utf-8', 'cp932', 'shift_jis', 'utf-16', 'latin1']:
            try:
                df = pd.read_csv(csv_file, encoding=encoding)
                print(f"エンコーディングに成功: {encoding}")
                break
            except UnicodeDecodeError:
                continue
        else:
            raise Exception("エンコーディングの自動判定に失敗しました。")

        tsv_file = f"{file_name}.tsv"
        df.to_csv(tsv_file, sep='\t', index=False)
        print(f"TSVファイルが作成されました: {tsv_file}")
    except Exception as e:
        print(f"エラーが発生しました: {e}")

if __name__ == "__main__":
    # ドラッグ＆ドロップされたファイルの処理
    if len(sys.argv) > 1:
        csv_file = sys.argv[1]  # ドロップされたファイルのパス
        if os.path.isfile(csv_file) and csv_file.endswith('.csv'):
            convert_csv_to_tsv(csv_file)
        else:
            print("有効なCSVファイルをドロップしてください。")
    else:
        print("CSVファイルをこのスクリプトにドラッグ＆ドロップしてください。")
