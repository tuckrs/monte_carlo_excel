; Monte Carlo Excel Add-in Installer Script
; NSIS (Nullsoft Scriptable Install System) script

!include "MUI2.nsh"
!include "LogicLib.nsh"
!include "x64.nsh"

; General
Name "Monte Carlo Excel Add-in"
OutFile "MonteCarloExcelSetup.exe"
InstallDir "$PROGRAMFILES64\Monte Carlo Excel"
InstallDirRegKey HKCU "Software\Monte Carlo Excel" ""
RequestExecutionLevel admin

; Interface Settings
!define MUI_ABORTWARNING
!define MUI_ICON "${NSISDIR}\Contrib\Graphics\Icons\modern-install.ico"
!define MUI_UNICON "${NSISDIR}\Contrib\Graphics\Icons\modern-uninstall.ico"

; Pages
!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_LICENSE "..\LICENSE"
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH

!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES

; Languages
!insertmacro MUI_LANGUAGE "English"

; Variables
Var PythonInstalled
Var PythonPath

Function .onInit
    ; Check if Python is installed
    ReadRegStr $PythonPath HKCU "Software\Python\PythonCore\3.12\InstallPath" ""
    ${If} $PythonPath == ""
        StrCpy $PythonInstalled "0"
    ${Else}
        StrCpy $PythonInstalled "1"
    ${EndIf}
FunctionEnd

Section "Prerequisites" SEC00
    ${If} $PythonInstalled == "0"
        MessageBox MB_OK|MB_ICONINFORMATION "Python 3.12 is required but not installed. Please download and install Python 3.12 from python.org, then run this installer again.$\n$\nMake sure to check 'Add Python to PATH' during installation."
        ExecShell "open" "https://www.python.org/downloads/"
        Abort
    ${EndIf}
SectionEnd

Section "MainSection" SEC01
    SetOutPath "$INSTDIR"
    SetOverwrite ifnewer
    
    ; Install files
    File "..\src\excel\monte_carlo.xlam"
    File "install_excel_addin.ps1"
    File "install_dependencies.py"
    
    ; Install Python dependencies
    DetailPrint "Installing Python dependencies..."
    nsExec::ExecToLog 'python "$INSTDIR\install_dependencies.py"'
    Pop $0
    ${If} $0 != "0"
        MessageBox MB_OK|MB_ICONSTOP "Failed to install Python dependencies!"
        Abort
    ${EndIf}
    
    ; Install Excel add-in
    DetailPrint "Installing Excel add-in..."
    nsExec::ExecToLog 'powershell -ExecutionPolicy Bypass -File "$INSTDIR\install_excel_addin.ps1" -AddinPath "$INSTDIR\monte_carlo.xlam"'
    Pop $0
    ${If} $0 != "0"
        MessageBox MB_OK|MB_ICONSTOP "Failed to install Excel add-in!"
        Abort
    ${EndIf}
    
    ; Create uninstaller
    WriteUninstaller "$INSTDIR\Uninstall.exe"
    
    ; Create start menu shortcut
    CreateDirectory "$SMPROGRAMS\Monte Carlo Excel"
    CreateShortCut "$SMPROGRAMS\Monte Carlo Excel\Uninstall.lnk" "$INSTDIR\Uninstall.exe"
    
    ; Write registry keys
    WriteRegStr HKCU "Software\Monte Carlo Excel" "" $INSTDIR
    WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\Monte Carlo Excel" \
                     "DisplayName" "Monte Carlo Excel Add-in"
    WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\Monte Carlo Excel" \
                     "UninstallString" "$\"$INSTDIR\Uninstall.exe$\""
SectionEnd

Section "Uninstall"
    ; Remove Excel add-in
    nsExec::ExecToLog 'powershell -ExecutionPolicy Bypass -Command "Remove-Item -Path \"$env:APPDATA\Microsoft\AddIns\monte_carlo.xlam\" -Force"'
    
    ; Remove registry keys
    DeleteRegKey HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\Monte Carlo Excel"
    DeleteRegKey HKCU "Software\Monte Carlo Excel"
    DeleteRegKey HKCU "Software\Microsoft\Office\Excel\AddIns\MonteCarloExcel"
    
    ; Remove files and uninstaller
    Delete "$INSTDIR\Uninstall.exe"
    Delete "$INSTDIR\monte_carlo.xlam"
    Delete "$INSTDIR\install_excel_addin.ps1"
    Delete "$INSTDIR\install_dependencies.py"
    
    ; Remove directories
    RMDir "$SMPROGRAMS\Monte Carlo Excel"
    RMDir "$INSTDIR"
SectionEnd
