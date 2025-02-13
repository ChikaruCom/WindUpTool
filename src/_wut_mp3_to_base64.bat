:: file:./ 
:: _script_run.bat
chcp 65001
@echo off
setlocal

REM 仮想環境名とPythonのバージョンを設定
set "ENV_NAME=py3.12_wut"

REM 実行スクリプトの絶対パスを取得し、拡張子を .py に変更
set "SCRIPT_DIR=%~dp0"
set "SCRIPT_PATH=%SCRIPT_DIR%%~n0.py"

REM 仮想環境をアクティベートしてスクリプトを実行
call conda activate %ENV_NAME%
python "%SCRIPT_PATH%" %1

REM 仮想環境をデアクティベート
call conda deactivate

endlocal
