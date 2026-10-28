; Offer Word Studio in Open With and Default Apps. Do not write extension
; defaults or Explorer's UserChoice: the user chooses their default handler.

!define WORD_STUDIO_CAPABILITIES "Software\WordStudio\Capabilities"
!define WORD_STUDIO_APPLICATION "Software\Classes\Applications\${APP_EXECUTABLE_FILENAME}"

!macro WordStudioDocumentTypes OPERATION
  !insertmacro ${OPERATION} "wordstudio" "Word Studio Editing Copy"
  !insertmacro ${OPERATION} "docx" "Word Document"
  !insertmacro ${OPERATION} "hwp" "Hangul Document"
  !insertmacro ${OPERATION} "hwpx" "Hangul XML Document"
  !insertmacro ${OPERATION} "html" "HTML Document"
  !insertmacro ${OPERATION} "htm" "HTML Document"
  !insertmacro ${OPERATION} "txt" "Text Document"
!macroend

!macro WordStudioRegisterDocument EXT DESCRIPTION
  WriteRegStr SHELL_CONTEXT "Software\Classes\WordStudio.Document.${EXT}" "" "${DESCRIPTION}"
  WriteRegStr SHELL_CONTEXT "Software\Classes\WordStudio.Document.${EXT}\DefaultIcon" "" '"$INSTDIR\${APP_EXECUTABLE_FILENAME}",0'
  WriteRegStr SHELL_CONTEXT "Software\Classes\WordStudio.Document.${EXT}\Application" "ApplicationName" "Word Studio"
  WriteRegStr SHELL_CONTEXT "Software\Classes\WordStudio.Document.${EXT}\shell\open\command" "" '"$INSTDIR\${APP_EXECUTABLE_FILENAME}" "%1"'
  WriteRegNone SHELL_CONTEXT "Software\Classes\.${EXT}\OpenWithProgids" "WordStudio.Document.${EXT}"
  WriteRegStr SHELL_CONTEXT "${WORD_STUDIO_CAPABILITIES}\FileAssociations" ".${EXT}" "WordStudio.Document.${EXT}"
  WriteRegStr SHELL_CONTEXT "${WORD_STUDIO_APPLICATION}\SupportedTypes" ".${EXT}" ""
!macroend

!macro WordStudioUnregisterDocument EXT DESCRIPTION
  DeleteRegValue SHELL_CONTEXT "Software\Classes\.${EXT}\OpenWithProgids" "WordStudio.Document.${EXT}"
  ; These keys are shared. NSIS /ifempty checks subkeys, not values, and
  ; would erase another handler or extension default during an upgrade.
  DeleteRegKey SHELL_CONTEXT "Software\Classes\WordStudio.Document.${EXT}"
!macroend

!macro customInstall
  WriteRegStr SHELL_CONTEXT "${WORD_STUDIO_CAPABILITIES}" "ApplicationName" "Word Studio"
  WriteRegStr SHELL_CONTEXT "${WORD_STUDIO_CAPABILITIES}" "ApplicationDescription" "Open and edit Word, Hangul, HTML and text documents."
  WriteRegStr SHELL_CONTEXT "${WORD_STUDIO_CAPABILITIES}" "ApplicationIcon" '"$INSTDIR\${APP_EXECUTABLE_FILENAME}",0'
  WriteRegStr SHELL_CONTEXT "Software\RegisteredApplications" "Word Studio" "${WORD_STUDIO_CAPABILITIES}"
  WriteRegStr SHELL_CONTEXT "${WORD_STUDIO_APPLICATION}" "FriendlyAppName" "Word Studio"
  WriteRegStr SHELL_CONTEXT "${WORD_STUDIO_APPLICATION}\shell\open\command" "" '"$INSTDIR\${APP_EXECUTABLE_FILENAME}" "%1"'
  !insertmacro WordStudioDocumentTypes WordStudioRegisterDocument
  System::Call 'shell32::SHChangeNotify(i, i, i, i) v (0x08000000, 0, 0, 0)'
!macroend

!macro customUnInstall
  !insertmacro WordStudioDocumentTypes WordStudioUnregisterDocument
  DeleteRegValue SHELL_CONTEXT "Software\RegisteredApplications" "Word Studio"
  DeleteRegKey SHELL_CONTEXT "${WORD_STUDIO_CAPABILITIES}"
  DeleteRegKey SHELL_CONTEXT "${WORD_STUDIO_APPLICATION}"
  System::Call 'shell32::SHChangeNotify(i, i, i, i) v (0x08000000, 0, 0, 0)'
!macroend
