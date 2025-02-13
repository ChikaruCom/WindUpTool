import sys
import base64
import os

def encode_mp3_to_base64(file_path):
    if not os.path.isfile(file_path):
        print(f"指定されたパスはファイルではありません: {file_path}")
        return

    try:
        # MP3ファイルを読み込み、Base64にエンコード
        with open(file_path, 'rb') as mp3_file:
            mp3_bytes = mp3_file.read()
            mp3_base64 = base64.b64encode(mp3_bytes).decode('utf-8')

        # 出力ファイル名を決定 (元のファイル名に ".txt" を追加)
        output_file_path = file_path + ".txt"
        
        # Base64エンコードされたデータをテキストファイルに書き出す
        with open(output_file_path, 'w') as output_file:
            output_file.write(mp3_base64)
        
        print(f"Base64エンコード結果が次のファイルに保存されました: {output_file_path}")
    
    except Exception as e:
        print(f"エラーが発生しました: {e}")

# 実行ファイルにファイルがドラッグされたかを確認
if len(sys.argv) > 1:
    mp3_file_path = sys.argv[1]  # ドラッグ＆ドロップされたファイルパス
    encode_mp3_to_base64(mp3_file_path)
else:
    print("MP3ファイルをドラッグ＆ドロップしてください。")
