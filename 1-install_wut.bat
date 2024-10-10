chcp 65001
@echo off
setlocal

:: フォルダを作成
mkdir "%UserProfile%\WindUpTool"
mkdir "%UserProfile%\WindUpTool\karakuri"
mkdir "%UserProfile%\WindUpTool\assets"
mkdir "%UserProfile%\WindUpTool\output"

REM Minicondaのインストール先とインストーラーのパスを設定
set "MINICONDA_URL=https://repo.anaconda.com/miniconda/Miniconda3-latest-Windows-x86_64.exe"
set "MINICONDA_INSTALLER=%UserProfile%\WindUpTool\miniconda.exe"
set "MINICONDA_PATH=%UserProfile%\WindUpTool\Miniconda3"

if not exist "%MINICONDA_INSTALLER%" (
    curl -L %MINICONDA_URL% -o %MINICONDA_INSTALLER%
    if %ERRORLEVEL% neq 0 (
        echo ダウンロードに失敗しました。
        exit /b %ERRORLEVEL%
    )
)

REM Minicondaが既にインストールされているかチェック
if not exist "%MINICONDA_PATH%" (
    echo Installing Miniconda...
    "%MINICONDA_INSTALLER%" /InstallationType=JustMe /RegisterPython=1 /AddToPath=0 /S /D=%MINICONDA_PATH%
)

REM Minicondaのパスを一時的に環境変数に追加
set "PATH=%MINICONDA_PATH%\Scripts;%MINICONDA_PATH%\condabin;%PATH%"

REM 仮想環境名とPythonのバージョンを設定
set "ENV_NAME=py3.12_wut"
set "PYTHON_VERSION=3.12"

REM 仮想環境が既に存在するかチェック
conda env list | findstr %ENV_NAME%
if %errorlevel% neq 0 (
    echo Creating environment %ENV_NAME%...
    conda create -y --name %ENV_NAME% python=%PYTHON_VERSION% pip
)

REM 仮想環境をアクティベートしてパッケージをインストール
call conda.bat activate %ENV_NAME%
pip install pandas==2.2.2 PyPDF2==3.0.1 pywin32==306 psutil==6.0.0 pyperclip==1.9.0 openpyxl=3.1.5

REM Condaの初期化 (次回のシェルでcondaが使えるようにする)
conda init

REM スクリプトの実行
set "SCRIPT_DIR=%~dp0"
set "SCRIPT_PATH=%SCRIPT_DIR%%~n0.py"
python %SCRIPT_PATH%

REM 仮想環境をデアクティベート
conda.bat deactivate

endlocal
