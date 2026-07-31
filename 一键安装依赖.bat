@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion
title 网络药理学脚本 — 一键安装依赖

:: ============================================
::  日志文件（桌面，方便查找）
:: ============================================
set "LOG=%USERPROFILE%\Desktop\安装日志_%date:~0,4%%date:~5,2%%date:~8,2%_%time:~0,2%%time:~3,2%%time:~6,2%.log"
set "LOG=%LOG: =0%"
echo === 网络药理学脚本依赖安装日志 === > "%LOG%"
echo 时间: %date% %time% >> "%LOG%"
echo 用户: %USERNAME% >> "%LOG%"
echo. >> "%LOG%"

echo.
echo ╔══════════════════════════════════════════════════════╗
echo ║     网络药理学 / 毒理学 脚本依赖一键安装 v2.0       ║
echo ║     日志文件: 桌面\安装日志_*.log                    ║
echo ╚══════════════════════════════════════════════════════╝
echo.

:: ============================================
::  步骤 1/6：检测 Python（多路径 + 版本校验）
:: ============================================
echo [1/6] 检测 Python 环境...
echo [1/6] 检测 Python 环境 >> "%LOG%"

set "PYTHON="

:: 尝试 python3 → python → py -3 → 搜索常见安装路径
for %%c in (python3 python) do (
    where %%c >nul 2>&1
    if !errorlevel! equ 0 (
        for /f "delims=" %%v in ('%%c --version 2^>^&1') do set "PYVER=%%v"
        echo   检测到: !PYVER! ^(%%c^)
        echo   检测到: !PYVER! ^(%%c^) >> "%LOG%"
        set "PYTHON=%%c"
        goto :py_found
    )
)

:: py -3 启动器
where py >nul 2>&1
if !errorlevel! equ 0 (
    for /f "delims=" %%v in ('py -3 --version 2^>^&1') do set "PYVER=%%v"
    echo   检测到: !PYVER! ^(py -3^)
    echo   检测到: !PYVER! ^(py -3^) >> "%LOG%"
    set "PYTHON=py -3"
    goto :py_found
)

:: 未找到
echo   [错误] 未检测到 Python
echo   [错误] 未检测到 Python >> "%LOG%"
goto :fail_python

:py_found

:: 校验版本号 ≥ 3.9
for /f "tokens=2 delims= " %%v in ("!PYVER!") do set "VER=%%v"
for /f "tokens=1,2 delims=." %%a in ("!VER!") do (
    set "MAJOR=%%a"
    set "MINOR=%%b"
)
if !MAJOR! lss 3 (
    echo   [错误] Python 版本 !VER! 不满足要求（需 ≥ 3.9）
    echo   [错误] Python 版本 !VER! 不满足要求 >> "%LOG%"
    goto :fail_python
)
if !MAJOR! equ 3 if !MINOR! lss 9 (
    echo   [错误] Python 版本 !VER! 不满足要求（需 ≥ 3.9）
    echo   [错误] Python 版本 !VER! 不满足要求 >> "%LOG%"
    goto :fail_python
)
echo   √ Python !VER! 满足要求
echo   √ Python !VER! 满足要求 >> "%LOG%"

:: 检查 pip 是否可用
echo   检查 pip...
!PYTHON! -m pip --version >nul 2>&1
if !errorlevel! neq 0 (
    echo   [错误] pip 不可用，尝试修复...
    echo   [错误] pip 不可用 >> "%LOG%"
    !PYTHON! -m ensurepip --upgrade >nul 2>&1
    if !errorlevel! neq 0 goto :fail_pip
)
echo   √ pip 可用
echo.

:: ============================================
::  步骤 2/6：网络检测 + 选择镜像源
:: ============================================
echo [2/6] 检测网络与镜像源...
echo [2/6] 检测网络与镜像源 >> "%LOG%"

set "PIP_INDEX="
set "PIP_TRUSTED="

:: 测试官方 PyPI 可达性
echo   测试 PyPI 官方源...
!PYTHON! -m pip install --dry-run pip --index-url https://pypi.org/simple/ --quiet --disable-pip-version-check >nul 2>&1
if !errorlevel! equ 0 (
    echo   √ PyPI 官方源可达
    echo   √ PyPI 官方源可达 >> "%LOG%"
) else (
    echo   - 官方源不可达，切换清华镜像...
    echo   - 官方源不可达，切换清华镜像 >> "%LOG%"
    set "PIP_INDEX=--index-url https://pypi.tuna.tsinghua.edu.cn/simple"
    set "PIP_TRUSTED=--trusted-host pypi.tuna.tsinghua.edu.cn"
    
    :: 验证清华镜像
    !PYTHON! -m pip install --dry-run pip !PIP_INDEX! !PIP_TRUSTED! --quiet --disable-pip-version-check >nul 2>&1
    if !errorlevel! equ 0 (
        echo   √ 清华镜像可用
        echo   √ 清华镜像可用 >> "%LOG%"
    ) else (
        echo   - 清华镜像不可达，尝试阿里云镜像...
        echo   - 清华镜像不可达，尝试阿里云镜像 >> "%LOG%"
        set "PIP_INDEX=--index-url https://mirrors.aliyun.com/pypi/simple/"
        set "PIP_TRUSTED=--trusted-host mirrors.aliyun.com"
    )
)

:: 升级 pip（带镜像）
echo   升级 pip...
!PYTHON! -m pip install --upgrade pip !PIP_INDEX! !PIP_TRUSTED! --quiet --disable-pip-version-check >> "%LOG%" 2>&1
echo   √ pip 已升级
echo.

:: ============================================
::  步骤 3/6：磁盘空间检查
:: ============================================
echo [3/6] 检查磁盘空间...
echo [3/6] 检查磁盘空间 >> "%LOG%"

cd /d "%~dp0"
for /f "tokens=3" %%a in ('dir /-c ^| findstr "可用" 2^>nul') do set "FREE=%%a"
if defined FREE (
    :: 大概需要 2GB（包 + Chromium）
    if !FREE! lss 2000000000 (
        echo   [警告] 可用空间约 !FREE! 字节（建议 2GB+），仍将尝试安装
        echo   [警告] 可用空间不足 >> "%LOG%"
    ) else (
        echo   √ 可用空间充足
    )
) else (
    echo   - 无法检测空间，跳过
)
echo.

:: ============================================
::  步骤 4/6：创建虚拟环境（可选，强烈推荐）
:: ============================================
echo [4/6] 设置虚拟环境...
echo [4/6] 设置虚拟环境 >> "%LOG%"

set "VENV_DIR=%~dp0venv"
set "USE_VENV=0"

:: 已有虚拟环境则直接用
if exist "!VENV_DIR!\Scripts\python.exe" (
    echo   √ 检测到已有虚拟环境，直接使用
    echo   √ 复用虚拟环境: !VENV_DIR! >> "%LOG%"
    set "PYTHON=!VENV_DIR!\Scripts\python.exe"
    set "USE_VENV=1"
    goto :venv_done
)

:: 询问用户
echo.
echo   ┌─────────────────────────────────────────────────┐
echo   │  推荐创建独立虚拟环境，避免污染系统 Python       │
echo   │  虚拟环境将创建在脚本同目录下的 venv 文件夹      │
echo   └─────────────────────────────────────────────────┘
echo.
set /p "ASK_VENV=  创建虚拟环境？[Y/n]: "
if /i "!ASK_VENV!"=="n" (
    echo   - 跳过虚拟环境，使用系统 Python
    echo   - 用户跳过虚拟环境 >> "%LOG%"
    goto :venv_done
)

:: 创建虚拟环境
echo   正在创建虚拟环境...
!PYTHON! -m venv "!VENV_DIR!" >> "%LOG%" 2>&1
if !errorlevel! neq 0 (
    echo   [警告] 虚拟环境创建失败，使用系统 Python 继续
    echo   [警告] 虚拟环境创建失败 >> "%LOG%"
    goto :venv_done
)

set "PYTHON=!VENV_DIR!\Scripts\python.exe"
set "USE_VENV=1"
echo   √ 虚拟环境创建成功: !VENV_DIR!
echo   √ 虚拟环境创建成功 >> "%LOG%"

:: 在虚拟环境中升级 pip
!PYTHON! -m pip install --upgrade pip !PIP_INDEX! !PIP_TRUSTED! --quiet --disable-pip-version-check >> "%LOG%" 2>&1

:venv_done
echo.

:: ============================================
::  步骤 5/6：安装依赖（逐包安装 + 重试 + 日志）
:: ============================================
echo [5/6] 安装 Python 依赖包...
echo [5/6] 安装 Python 依赖包 >> "%LOG%"
echo   请耐心等待，每个包安装约 3-30 秒...
echo.

:: 包清单：名称 | 导入名 | 类别（core=核心 / web=浏览器 / chem=化学 / opt=可选）
:: 含安装提示语
set "PKG_COUNT=0"
set "FAIL_COUNT=0"
set "PKG_LIST="

call :install_pkg "requests"            "requests"            "core" "HTTP 请求库"
call :install_pkg "openpyxl"            "openpyxl"            "core" "Excel 文件读写"
call :install_pkg "pandas"              "pandas"              "core" "数据处理与表格"
call :install_pkg "beautifulsoup4"      "bs4"                 "core" "HTML 页面解析"
call :install_pkg "colorama"            "colorama"            "core" "终端彩色输出"
call :install_pkg "playwright"          "playwright.sync_api" "web"  "浏览器自动化(核心)"
call :install_pkg "selenium"            "selenium"            "web"  "浏览器自动化(备选)"
call :install_pkg "webdriver-manager"   "webdriver_manager"   "web"  "Chrome 驱动管理"

:: rdkit 是出名的 Windows 安装困难户，优先尝试 conda 路径，再 pip
call :install_rdkit

call :install_pkg "openbabel-wheel"     "openbabel"           "opt"  "MOL2 格式输出(可选)"

echo.
echo   安装统计: 共 !PKG_COUNT! 个包，失败 !FAIL_COUNT! 个
echo   安装统计: 共 !PKG_COUNT! 个包，失败 !FAIL_COUNT! 个 >> "%LOG%"
echo.

if !FAIL_COUNT! gtr 0 (
    echo   [注意] 部分包安装失败，详见下方标记 × 的项
    echo   [注意] 部分包安装失败 >> "%LOG%"
)
echo.

:: ============================================
::  步骤 6/6：安装 Playwright Chromium + 验证
:: ============================================
echo [6/6] Playwright Chromium 浏览器 + 最终验证...
echo [6/6] Playwright Chromium 浏览器 + 最终验证 >> "%LOG%"

:: 检查 playwright 是否已装（不装这个 Chromium 就没意义）
!PYTHON! -c "from playwright.sync_api import sync_playwright" >nul 2>&1
if !errorlevel! equ 0 (
    echo   安装 Chromium 浏览器 ^(约 150MB，网速慢需 2-5 分钟^)...
    echo   安装 Playwright Chromium >> "%LOG%"
    
    !PYTHON! -m playwright install chromium >> "%LOG%" 2>&1
    if !errorlevel! neq 0 (
        :: 重试一次
        echo   [重试] Chromium 安装失败，再次尝试...
        !PYTHON! -m playwright install chromium >> "%LOG%" 2>&1
        if !errorlevel! neq 0 (
            echo   × Chromium 未安装 ^(不影响其他脚本^)
            echo   × Chromium 未安装 >> "%LOG%"
            echo   手动安装: !PYTHON! -m playwright install chromium
        ) else (
            echo   √ Chromium 安装成功
            echo   √ Chromium 安装成功 >> "%LOG%"
        )
    ) else (
        echo   √ Chromium 安装成功
        echo   √ Chromium 安装成功 >> "%LOG%"
    )
) else (
    echo   - 跳过 Chromium（playwright 未安装）
    echo   - 跳过 Chromium >> "%LOG%"
)
echo.

:: ── 最终验证 ──
echo ══════════════════════════════════════════════════════
echo   最终验证（可导入 = √，失败 = ×）
echo ══════════════════════════════════════════════════════
echo.
echo   最终验证 >> "%LOG%"

set "ALL_OK=1"

call :verify "requests"           "HTTP 请求"
call :verify "openpyxl"           "Excel 读写"
call :verify "pandas"             "数据处理"
call :verify "bs4"                "HTML 解析"
call :verify "colorama"           "终端彩色"
call :verify_import "from rdkit import Chem"           "RDKit 化学计算"
call :verify_import "from playwright.sync_api import sync_playwright" "Playwright 自动化"
call :verify "selenium"          "Selenium 自动化"
call :verify_import "from openbabel import pybel"      "MOL2 格式支持(可选)"

echo.
echo ══════════════════════════════════════════════════════
echo   最终验证 >> "%LOG%"

:: ============================================
::  完成
:: ============================================
echo.
echo ╔══════════════════════════════════════════════════════╗
if !ALL_OK! equ 1 (
    echo ║            全部依赖安装成功！                      ║
) else (
    echo ║          安装完成（部分可选包未成功）               ║
)
echo ╚══════════════════════════════════════════════════════╝
echo.
if "!USE_VENV!"=="1" (
    echo   ✔ 虚拟环境: !VENV_DIR!
)
echo   ✔ 日志文件: %LOG%
echo.
echo   ── 脚本清单 ──
echo   ① TTP.py                     TTD 疾病靶点搜索
echo   ② chembl_target_search_*.py  ChEMBL 化合物靶点检索
echo   ③ disgenet.py                DisGeNET 疾病靶点检索
echo   ④ open_targets.py            Open Targets 疾病靶点检索
echo   ⑤ pharmgkb.py                PharmGKB 基因靶点搜索
echo   ⑥ superpred.py               SuperPred AI 靶点预测
echo   ⑦ swiss_target_prediction_*.py  SwissTargetPrediction
echo   ⑧ targetnet_interactive.py   TargetNet QSAR 靶点预测
echo   ⑨ 计算化学_小分子处理.py     分子属性 & 结构处理
echo.
echo   启动: python 脚本名.py
echo.
pause
exit /b 0

:: ============================================
::  失败退出
:: ============================================
:fail_python
echo.
echo ╔══════════════════════════════════════════════════════╗
echo ║  安装失败：未检测到 Python 3.9+                     ║
echo ╚══════════════════════════════════════════════════════╝
echo.
echo   请按以下步骤操作：
echo   1. 访问 https://www.python.org/downloads/
echo   2. 下载 Python 3.9 及以上版本
echo   3. 安装时务必勾选 "Add Python to PATH"
echo   4. 重新运行本安装脚本
echo.
echo   日志文件: %LOG%
pause
exit /b 1

:fail_pip
echo.
echo   [错误] pip 不可用且无法自动修复
echo   请手动执行: !PYTHON! -m ensurepip --upgrade
echo   日志文件: %LOG%
pause
exit /b 1

:: ============================================
::  子函数：安装单个包（带重试）
:: ============================================
:install_pkg
:: %1=pkg名  %2=导入名  %3=类别  %4=说明
set /a PKG_COUNT+=1

:: 先检查是否已安装
!PYTHON! -c "import %~2" >nul 2>&1
if !errorlevel! equ 0 (
    echo   √ %~1 ^(%~4^) [已安装]
    echo   √ %~1 已安装 >> "%LOG%"
    goto :eof
)

echo   → 安装 %~1 ^(%~4^)...
echo   → 安装 %~1 >> "%LOG%"

:: 第 1 次尝试
!PYTHON! -m pip install %~1 !PIP_INDEX! !PIP_TRUSTED! --disable-pip-version-check >> "%LOG%" 2>&1
if !errorlevel! equ 0 (
    call :verify_pkg "%~1" "%~2" "%~4"
    goto :eof
)

:: 第 2 次尝试（使用 --user）
echo     重试(2/3) --user 模式...
!PYTHON! -m pip install %~1 --user !PIP_INDEX! !PIP_TRUSTED! --disable-pip-version-check >> "%LOG%" 2>&1
if !errorlevel! equ 0 (
    call :verify_pkg "%~1" "%~2" "%~4"
    goto :eof
)

:: 第 3 次尝试（--no-cache-dir）
echo     重试(3/3) --no-cache-dir...
!PYTHON! -m pip install %~1 --no-cache-dir !PIP_INDEX! !PIP_TRUSTED! --disable-pip-version-check >> "%LOG%" 2>&1
if !errorlevel! equ 0 (
    call :verify_pkg "%~1" "%~2" "%~4"
    goto :eof
)

:: 3 次全失败
set /a FAIL_COUNT+=1
echo   × %~1 ^(%~4^) [安装失败]
echo   × %~1 安装失败 >> "%LOG%"
goto :eof

:: ── 子函数：安装后导入验证 ──
:verify_pkg
!PYTHON! -c "import %~2" >nul 2>&1
if !errorlevel! equ 0 (
    echo   √ %~1 ^(%~3^)
    echo   √ %~1 验证通过 >> "%LOG%"
) else (
    set /a FAIL_COUNT+=1
    echo   × %~1 ^(%~3^) [导入失败]
    echo   × %~1 导入失败 >> "%LOG%"
)
goto :eof

:: ============================================
::  子函数：安装 rdkit（特殊处理）
:: ============================================
:install_rdkit
set /a PKG_COUNT+=1

!PYTHON! -c "from rdkit import Chem" >nul 2>&1
if !errorlevel! equ 0 (
    echo   √ rdkit ^(化学计算^) [已安装]
    echo   √ rdkit 已安装 >> "%LOG%"
    goto :eof
)

echo   → 安装 rdkit ^(化学计算，可能较慢^)...
echo   → 安装 rdkit >> "%LOG%"

:: 尝试 pip 安装（新版 rdkit 对 Windows 兼容好了很多）
!PYTHON! -m pip install rdkit !PIP_INDEX! !PIP_TRUSTED! --disable-pip-version-check >> "%LOG%" 2>&1
if !errorlevel! equ 0 (
    call :verify_pkg "rdkit" "rdkit" "化学计算"
    goto :eof
)

:: 尝试 conda（如果装了）
where conda >nul 2>&1
if !errorlevel! equ 0 (
    echo     检测到 conda，尝试 conda 安装...
    echo     尝试 conda install rdkit >> "%LOG%"
    conda install -y -c conda-forge rdkit >> "%LOG%" 2>&1
    if !errorlevel! equ 0 (
        call :verify_pkg "rdkit" "rdkit" "化学计算(conda)"
        goto :eof
    )
)

:: rdkit 装不上给出明确指引
set /a FAIL_COUNT+=1
echo   × rdkit ^(化学计算^) [安装失败]
echo   × rdkit 安装失败 >> "%LOG%"
echo      ┌─────────────────────────────────────────────┐
echo      │ rdkit 安装失败，可选择以下方案：            │
echo      │                                            │
echo      │ A. 安装 Anaconda 后用 conda 安装 rdkit：    │
echo      │    conda install -c conda-forge rdkit       │
echo      │                                            │
echo      │ B. 仅影响「计算化学_小分子处理.py」        │
echo      │    其他 8 个脚本不受影响                    │
echo      └─────────────────────────────────────────────┘
goto :eof

:: ============================================
::  子函数：验证导入
:: ============================================
:verify
:: %1=导入名  %2=说明
!PYTHON! -c "import %~1" >nul 2>&1
if !errorlevel! equ 0 (
    echo   √ %~1 ^(%~2^)
    echo   √ %~1 >> "%LOG%"
) else (
    echo   × %~1 ^(%~2^)
    echo   × %~1 >> "%LOG%"
    set "ALL_OK=0"
)
goto :eof

:verify_import
:: %1=完整import语句  %2=说明
!PYTHON! -c "%~1" >nul 2>&1
if !errorlevel! equ 0 (
    echo   √ %~2
    echo   √ %~2 >> "%LOG%"
) else (
    echo   × %~2
    echo   × %~2 >> "%LOG%"
    set "ALL_OK=0"
)
goto :eof
