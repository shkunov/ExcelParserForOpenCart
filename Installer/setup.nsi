;Include Modern UI

  !define MUI_VERSION "1.13"
  !define MUI_PRODUCT "ExcelParserForOpenCart"
  !define MUI_FILE "ExcelParserForOpenCart"
  !define MUI_ICON "program.ico"
  !define MUI_UNICON "program.ico"
  !include "MUI2.nsh"

;--------------------------------
;General

  ;Name and file
  Name "ExcelParserForOpenCart-${MUI_VERSION}"
  OutFile "ExcelParserForOpenCart-${MUI_VERSION}.exe"
  SetCompressor /FINAL /SOLID lzma

  ;Default installation folder
  InstallDir "$PROGRAMFILES\ExcelParserForOpenCart"
  
  ;Get installation folder from registry if available
  InstallDirRegKey HKCU "Software\ExcelParserForOpenCart" ""

  ;Request application privileges for Windows Vista
  RequestExecutionLevel user

;--------------------------------
;Interface Settings

  !define MUI_ABORTWARNING

;--------------------------------
;Pages
  !insertmacro MUI_PAGE_DIRECTORY
  !insertmacro MUI_PAGE_INSTFILES
  # These indented statements modify settings for MUI_PAGE_FINISH
    !define MUI_FINISHPAGE_NOAUTOCLOSE
    !define MUI_FINISHPAGE_RUN
    !define MUI_FINISHPAGE_RUN_NOTCHECKED
    !define MUI_FINISHPAGE_RUN_TEXT "Start program"
    !define MUI_FINISHPAGE_RUN_FUNCTION "LaunchLink"
    !define MUI_FINISHPAGE_SHOWREADME_NOTCHECKED
    !define MUI_FINISHPAGE_SHOWREADME $INSTDIR\readme.txt
  !insertmacro MUI_PAGE_FINISH
  
  !insertmacro MUI_UNPAGE_CONFIRM
  !insertmacro MUI_UNPAGE_INSTFILES
  
;--------------------------------
;Languages
 
  !insertmacro MUI_LANGUAGE "Russian"

;--------------------------------
;Installer Sections

Section "Dummy Section" SecDummy

  SetOutPath "$INSTDIR"
  File "Readme.txt"
  File "History.txt"
  File "..\Output\Release\${MUI_FILE}.exe"
  File "..\Output\Release\*.*"
  SetOutPath "$INSTDIR\x86"
  File "..\Output\Release\x86\*.*"
  SetOutPath "$INSTDIR\x64"
  File "..\Output\Release\x64\*.*"

  CreateShortCut "$DESKTOP\${MUI_PRODUCT}.lnk" "$INSTDIR\${MUI_FILE}.exe" ""
 
  ;create start-menu items
  CreateDirectory "$SMPROGRAMS\${MUI_PRODUCT}"
  CreateShortCut "$SMPROGRAMS\${MUI_PRODUCT}\Uninstall.lnk" "$INSTDIR\Uninstall.exe" "" "$INSTDIR\Uninstall.exe" 0
  CreateShortCut "$SMPROGRAMS\${MUI_PRODUCT}\${MUI_PRODUCT}.lnk" "$INSTDIR\${MUI_FILE}.exe" "" "$INSTDIR\${MUI_FILE}.exe" 0
  CreateShortCut "$SMPROGRAMS\${MUI_PRODUCT}\Readme.lnk" "$INSTDIR\Readme.txt" "" "$INSTDIR\Readme.txt" 0
  CreateShortCut "$SMPROGRAMS\${MUI_PRODUCT}\History.lnk" "$INSTDIR\History.txt" "" "$INSTDIR\History.txt" 0
  
  ;Store installation folder
  WriteRegStr HKCU "Software\ExcelParserForOpenCart" "" $INSTDIR
  
  ;Create uninstaller
  WriteUninstaller "$INSTDIR\Uninstall.exe"

SectionEnd

;--------------------------------
;Descriptions

  ;Language strings
  LangString DESC_SecDummy ${LANG_ENGLISH} "A test section."

  ;Assign language strings to sections
  !insertmacro MUI_FUNCTION_DESCRIPTION_BEGIN
    !insertmacro MUI_DESCRIPTION_TEXT ${SecDummy} $(DESC_SecDummy)
  !insertmacro MUI_FUNCTION_DESCRIPTION_END

;--------------------------------
;Uninstaller Section

Section "Uninstall"

  Delete "$INSTDIR\*.*"
  Delete "$INSTDIR\x86\*.*"
  RMDir "$INSTDIR\x86"
  Delete "$INSTDIR\x64\*.*"
  RMDir "$INSTDIR\x64"
  Delete "$INSTDIR\Uninstall.exe"

  RMDir "$INSTDIR"

  ;Delete Start Menu Shortcuts
  Delete "$DESKTOP\${MUI_PRODUCT}.lnk"
  Delete "$SMPROGRAMS\${MUI_PRODUCT}\*.*"
  RmDir  "$SMPROGRAMS\${MUI_PRODUCT}"

  DeleteRegKey /ifempty HKCU "Software\ExcelParserForOpenCart"

SectionEnd

Function LaunchLink
  ;MessageBox MB_OK "Reached LaunchLink $\r$\n \
  ;                 SMPROGRAMS: $SMPROGRAMS  $\r$\n \
  ;                 Start Menu Folder: $STARTMENU_FOLDER $\r$\n \
  ;                 InstallDirectory: $INSTDIR "
  ExecShell "" "$INSTDIR\${MUI_FILE}.exe"
FunctionEnd