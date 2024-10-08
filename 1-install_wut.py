import subprocess
import sys
import importlib
import os
import shutil
from glob import glob
import getpass

# スクリプトが実行されているディレクトリの絶対パスを取得
script_dir = os.path.dirname(os.path.abspath(__file__))

def execute_registry_script(registry_file):
    try:
        # .regファイルをregeditを使って実行
        subprocess.run(["regedit", "/s", registry_file], check=True)
        print(f"{registry_file} をレジストリに適用しました。")
    except subprocess.CalledProcessError as e:
        print(f"レジストリの適用中にエラーが発生しました: {e}")


def generate_registry_script():
    # 実行時のユーザー名を取得
    user_name = getpass.getuser()

    # 基本設定
    top_menu_name = "WindUpTool"
    icon_path = f"C:\\Users\\{user_name}\\WindUpTool\\assets\\ico_wut.ico"
    bat_files_path = f"C:\\Users\\{user_name}\\WindUpTool\\karakuri\\*.bat"
    
    # バッチファイルのパスを取得
    bat_files = glob(bat_files_path)
    

    # 対象セグメント
    segments = [f"*", f"Directory\\Background"]

    # レジストリデータ
    registry_data = []

    for target in segments:
        # トップメニューの設定
        # 変数展開考慮
        escape_target = target.replace("\\", "\\\\")
        registry_data.append(f'[HKEY_CLASSES_ROOT\\{target}\\shell\\{top_menu_name}]')
        registry_data.append(f'"MUIVerb"="Wind Up Tool"')
        registry_data.append(f'"ExtendedSubCommandsKey"="{escape_target}\\\\shell\\\\{top_menu_name}\\\\submenu1"')
        registry_data.append(f'"Icon"="\\"C:\\\\Users\\\\{user_name}\\\\WindUpTool\\\\assets\\\\ico_wut.ico\\""')

        # サブメニューの設定
        registry_data.append(f'')
        registry_data.append(f'[HKEY_CLASSES_ROOT\\{target}\\shell\\{top_menu_name}\\submenu1]')
        registry_data.append(f'[HKEY_CLASSES_ROOT\\{target}\\shell\\{top_menu_name}\\submenu1\\shell]')
        registry_data.append(f'')

        # サブメニューの追加
        for bat_file in bat_files:
            file_name = os.path.basename(bat_file)
            menu_name = os.path.splitext(file_name)[0]

            # "script表記がどちらも有効"
            if "script" not in file_name:
                # データタイプの場合はdirスクリプトを除く
                if "*" in target and "dir" in file_name:
                    continue
                # dirタイプの場合はdirスクリプトのみ
                if "Directory" in target and "dir" not in file_name:
                    continue

            # Directoryタイプの引数は%V
            if "Directory" in target:
                registry_data.append(f'[HKEY_CLASSES_ROOT\\{target}\\shell\\{top_menu_name}\\submenu1\\shell\\{menu_name}]')
                registry_data.append(f'@="{menu_name}"')
                registry_data.append(f'')
                registry_data.append(f'[HKEY_CLASSES_ROOT\\{target}\\shell\\{top_menu_name}\\submenu1\\shell\\{menu_name}\\command]')
                registry_data.append(f'@="C:\\\\Users\\\\{user_name}\\\\WindUpTool\\\\karakuri\\\\{file_name} \\"%V\\""')
                registry_data.append(f'')
                registry_data.append(f'')
            else:
                registry_data.append(f'[HKEY_CLASSES_ROOT\\{target}\\shell\\{top_menu_name}\\submenu1\\shell\\{menu_name}]')
                registry_data.append(f'@="{menu_name}"')
                registry_data.append(f'')
                registry_data.append(f'[HKEY_CLASSES_ROOT\\{target}\\shell\\{top_menu_name}\\submenu1\\shell\\{menu_name}\\command]')
                registry_data.append(f'@="C:\\\\Users\\\\{user_name}\\\\WindUpTool\\\\karakuri\\\\{file_name} \\"%1\\""')
                registry_data.append(f'')
                registry_data.append(f'')

    # レジストリデータを一つの文字列にまとめる
    registry_script = '\n'.join(registry_data)
    
    # レジストリファイルの保存
    output_file = f"WindUpTool_registry_{user_name}.reg"
    with open(output_file, 'w') as f:
        f.write("Windows Registry Editor Version 5.00\n\n")
        f.write(registry_script)
    
    print(f"レジストリファイルが生成されました: {output_file}")

    execute_registry_script(output_file)

def print_interpreter_path():
    interpreter_path = sys.executable
    print(f"Python interpreter path: {interpreter_path}")

def install_pip():
    try:
        import pip
    except ImportError:
        print("pip not found. Installing pip...")
        subprocess.check_call([sys.executable, '-m', 'ensurepip', '--upgrade'])
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', '--upgrade', 'pip'])
        print("pip installed successfully.")

def install_and_import(module_name, version=None):
    try:
        if version:
            module = importlib.import_module(module_name)
            installed_version = module.__version__
            if installed_version != version:
                raise ImportError(f"Version mismatch: {installed_version} != {version}")
        else:
            importlib.import_module(module_name)
    except ImportError:
        if version:
            package = f"{module_name}=={version}"
        else:
            package = module_name
        print(f"{package} not found or version mismatch. Installing...")
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', package])
        print(f"{package} installed successfully.")
        importlib.import_module(module_name)

def main():
    # pipがインストールされているか確認し、インストール
    install_pip()

    # チェックするモジュールとバージョンのリスト
    modules_to_check = [
#        ("numpy", None),  # numpy
        ("pandas", None),  # pandas
        ("PyPDF2", None),  # PyPDF2
        ("win32com", None),  # win32com
        ("psutil", None),
        ("pyperclip", None),
#        ("requests", None)  # requests
    ]

    for module, version in modules_to_check:
        install_and_import(module, version)

    # Pythonインタープリタのパスを表示
    print(f"Python interpreter path: {sys.executable}")

    # src -> karakuri
    # コピー元のパス
    src_folder = os.path.join(script_dir, 'src', '*')

    # 環境変数 %UserProfile% を取得して、コピー先のディレクトリを設定
    dst_folder = os.path.join(os.environ['USERPROFILE'], 'WindUpTool', 'karakuri')

    # コピー先フォルダが存在しない場合は作成する
    os.makedirs(dst_folder, exist_ok=True)

    # srcフォルダ内の全ファイルを取得してコピー
    for file in glob(src_folder):
        if os.path.isfile(file):
            shutil.copy(file, dst_folder)
            print(f"Copied {file} to {dst_folder}")
        else:
            print(f"Skipped {file}, not a file")

    # 環境変数 %UserProfile% を取得して、コピー先のディレクトリを設定
    dst_folder = os.path.join(os.environ['USERPROFILE'], 'WindUpTool', 'assets')

    # アイコンの移動
    ico_path = os.path.join(script_dir, 'ico', 'ico_wut.ico')
    shutil.copy(ico_path, dst_folder)

if __name__ == "__main__":
    main()
    generate_registry_script()
    print("your_script: OK")
    print_interpreter_path()
