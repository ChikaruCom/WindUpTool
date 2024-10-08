import os
import sys
from PIL import Image

def create_multisize_icon(png_path):
    # PNGファイルを開く
    img = Image.open(png_path)

    # 元のファイル名（拡張子なし）
    base_name = os.path.splitext(os.path.basename(png_path))[0]

    # アイコンに含めるサイズのリスト
    icon_sizes = [(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]

    # マルチサイズアイコンの保存
    output_ico_path = f"{base_name}.ico"
    icon_images = [img.resize(size, Image.ANTIALIAS) for size in icon_sizes]
    img.save(output_ico_path, format='ICO', sizes=icon_sizes)
    print(f"マルチサイズアイコン {output_ico_path} が作成されました。")

    # 各サイズごとのICOファイルを保存
    for size in icon_sizes:
        resized_img = img.resize(size, Image.ANTIALIAS)
        individual_icon_path = f"{base_name}_{size[0]}x{size[1]}.ico"
        resized_img.save(individual_icon_path, format='ICO')
        print(f"{size[0]}x{size[1]} のアイコン {individual_icon_path} が作成されました。")

def main():
    # ドラッグ&ドロップされたPNGファイルを取得
    if len(sys.argv) > 1:
        png_path = sys.argv[1]

        # PNGファイルが存在するか確認
        if os.path.exists(png_path) and png_path.lower().endswith('.png'):
            create_multisize_icon(png_path)
        else:
            print("指定されたファイルが存在しないか、PNGファイルではありません。")
    else:
        print("PNGファイルをドラッグ&ドロップしてください。")

if __name__ == "__main__":
    main()
