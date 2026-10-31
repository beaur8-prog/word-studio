!include "${__FILEDIR__}\windows-file-associations.nsh"

Var WordStudioExistingInstall
Var WordStudioMachineInstall

!macro preInit
  ; Windows x64 installers write their registration in the 64-bit view.
  SetRegView 64
  StrCpy $WordStudioExistingInstall "0"
  StrCpy $WordStudioMachineInstall ""
  ReadRegStr $R0 HKCU "${INSTALL_REGISTRY_KEY}" InstallLocation
  ${If} ${FileExists} "$R0\${APP_EXECUTABLE_FILENAME}"
    StrCpy $WordStudioExistingInstall "1"
  ${EndIf}
  ReadRegStr $R0 HKLM "${INSTALL_REGISTRY_KEY}" InstallLocation
  ${If} ${FileExists} "$R0\${APP_EXECUTABLE_FILENAME}"
    StrCpy $WordStudioExistingInstall "1"
    StrCpy $WordStudioMachineInstall $R0
  ${EndIf}
!macroend

!macro customInstallMode
  ; Older helpers supplied /currentuser even for a machine installation.
  ; With no explicit target folder, repair that existing installation instead
  ; of creating a second copy under LocalAppData. Standard NSIS upgrade removal
  ; also reconciles an existing per-user copy while preserving app data.
  !insertmacro GetDParameter $R0
  ${If} $R0 == ""
  ${AndIf} $WordStudioMachineInstall != ""
    StrCpy $isForceMachineInstall "1"
  ${EndIf}
!macroend

!macro customFinishPage
  !define MUI_FINISHPAGE_RUN
  !define MUI_FINISHPAGE_RUN_FUNCTION WordStudioFinishRun
  !define MUI_PAGE_CUSTOMFUNCTION_SHOW WordStudioFinishShow
  !insertmacro MUI_PAGE_FINISH

  Function WordStudioFinishRun
    ; A fresh installation may launch from its finish page. On reinstall or
    ; update, only the waiting helper may launch after it verifies the version.
    ${If} $WordStudioExistingInstall == "0"
      ${StdUtils.ExecShellAsUser} $0 "$launchLink" "open" ""
    ${EndIf}
  FunctionEnd

  Function WordStudioFinishShow
    ${If} $WordStudioExistingInstall == "1"
      SendMessage $mui.FinishPage.Run ${BM_SETCHECK} ${BST_UNCHECKED} 0
      ShowWindow $mui.FinishPage.Run ${SW_HIDE}
    ${EndIf}
  FunctionEnd

!macroend
