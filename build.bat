@echo off
set "vswhere=%ProgramFiles(x86)%\Microsoft Visual Studio\Installer\vswhere.exe"

echo Attempting to locate Visual Studio...

if not exist "%vswhere%" (
    echo vswhere.exe not found. Visual Studio Installer might not be installed.
    echo Please install Visual Studio with C++ workload.
    pause
    exit /b 1
)

echo.
echo Cleaning old build files...
if exist resource.res del resource.res
if exist media_sorter.obj del media_sorter.obj
if exist sorter.obj del sorter.obj
if exist "Media Sorter XXL.exe" del "Media Sorter XXL.exe"

for /f "usebackq tokens=*" %%i in (`"%vswhere%" -latest -products * -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 -property installationPath`) do (
    set "VSInstallDir=%%i"
)

if not "%VSInstallDir:~-1%"=="\" set "VSInstallDir=%VSInstallDir%\"

if not defined VSInstallDir (
    echo Visual Studio with C++ Tools not found.
    echo Please ensure you have installed the "Desktop development with C++" workload.
    pause
    exit /b 1
)

echo Found Visual Studio at: "%VSInstallDir%"

if exist "%VSInstallDir%\VC\Auxiliary\Build\vcvars64.bat" (
    call "%VSInstallDir%\VC\Auxiliary\Build\vcvars64.bat"
) else (
    echo Error: vcvars64.bat not found in expected location.
    pause
    exit /b 1
)

echo.
echo Compiling Resources...
rc.exe /nologo resource.rc

echo.
echo Compiling Sorter...
cl.exe /nologo /O2 /EHsc /std:c++17 /DUNICODE /D_UNICODE /utf-8 media_sorter.cpp resource.res /link /SUBSYSTEM:WINDOWS /OUT:"Media Sorter XXL.exe"

if %errorlevel% neq 0 (
    echo Compilation Failed!
    pause
    exit /b %errorlevel%
)

echo.
echo Compilation Success! "Media Sorter XXL.exe" created.
echo.
pause
