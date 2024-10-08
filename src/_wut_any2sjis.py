# 推奨されるスクリプト名: convert_to_sjis.py: any2sjis
import os
import sys
import chardet

def detect_encoding(file_path):
    # エンコードを自動検出
    with open(file_path, 'rb') as f:
        raw_data = f.read()
    result = chardet.detect(raw_data)
    return result['encoding']

def convert_to_sjis(file_path):
    # ファイルの拡張子を取得
    file_ext = os.path.splitext(file_path)[1].lower()

    # サポートされているファイル形式かを確認
    if file_ext not in ['.txt', '.csv', '.tsv']:
        print(f"Unsupported file format: {file_ext}")
        return

    # ファイルのエンコードを検出
    encoding = detect_encoding(file_path)
    print(f"Detected encoding for {file_path}: {encoding}")

    # 出力ファイル名を設定（_sjisを追加）
    base_name, ext = os.path.splitext(file_path)
    new_file_path = f"{base_name}_sjis{ext}"

    try:
        # 元ファイルを適切なエンコードで読み込み
        with open(file_path, 'r', encoding=encoding) as f:
            content = f.read()

        # Shift_JISでファイルを書き出し
        with open(new_file_path, 'w', encoding='shift_jis') as f:
            f.write(content)

        print(f"File converted and saved as: {new_file_path}")

    except Exception as e:
        print(f"Error processing file {file_path}: {e}")

if __name__ == "__main__":
    # D&Dされたファイルを処理
    if len(sys.argv) > 1:
        for file in sys.argv[1:]:
            convert_to_sjis(file)
    else:
        print("Please drag and drop the file(s) onto the script.")
