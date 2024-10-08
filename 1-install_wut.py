import subprocess
import sys
import importlib

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

if __name__ == "__main__":
    main()
    print("your_script: OK")
    print_interpreter_path()
