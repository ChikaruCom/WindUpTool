# 推奨されるスクリプト名: convert_csv_to_tsv.py
# このスクリプトは、CSVファイルをドラッグ＆ドロップすることでTSV形式に変換します。
# CSVファイルをスクリプトにドロップすると、同じディレクトリに同名のTSVファイルが生成されます。

import pandas as pd
import sys
import os
import whisper
import json

# モデルの読み出し（今回はsmallモデルを利用）
#type = "small"
type = "medium"
#type = "large"

#model = whisper.load_model("small")
#model = whisper.load_model("large")
model = whisper.load_model(type)


def convert_csv_to_tsv(csv_file):
    
    result = model.transcribe(csv_file, verbose=True, language='ja')
    print(result['text'])

# resultをJSON形式で保存する
    with open('transcription_result.json', 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=4)

if __name__ == "__main__":
    # ドラッグ＆ドロップされたファイルの処理
    if len(sys.argv) > 1:
        csv_file = sys.argv[1]  # ドロップされたファイルのパス
        convert_csv_to_tsv(csv_file)
    else:
        print("CSVファイルをこのスクリプトにドラッグ＆ドロップしてください。")
