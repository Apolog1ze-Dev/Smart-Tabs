; Smart Tabs Context Menu Launcher
; This is a non-elevated launcher that handles context menu requests
; and communicates with the main elevated Smart Tabs process

#NoTrayIcon
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=..\build\icon.ico
#AutoIt3Wrapper_Outfile=smart-tabs-launcher.exe
#AutoIt3Wrapper_UseX64=y
#AutoIt3Wrapper_Res_Comment=Smart Tabs Context Menu Launcher
#AutoIt3Wrapper_Res_Description=Non-elevated launcher for Smart Tabs context menu integration
#AutoIt3Wrapper_Res_Fileversion=1.0.0.0
#AutoIt3Wrapper_Res_ProductName=Smart Tabs Context Launcher
#AutoIt3Wrapper_Res_ProductVersion=1.0
#AutoIt3Wrapper_Res_CompanyName=Smart Tabs
#AutoIt3Wrapper_Res_LegalCopyright=Copyright Smart Tabs
; Note: NO requireAdministrator - this runs with normal user privileges
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

; Constants
Global Const $SMART_TABS_PIPE_NAME = "\\.\pipe\SmartTabsContextMenu"
Global Const $SMART_TABS_EXE_NAME = "Smart Tabs.exe"
Global Const $MAX_WAIT_TIME = 5000 ; 5 seconds

; Main execution
Main()

Func Main()
    ; Parse command line arguments
    If $CmdLine[0] < 1 Then
        ConsoleWrite("Usage: smart-tabs-launcher.exe [--add-shortcut <file-path> | --autostart]" & @CRLF)
        Exit(1)
    EndIf
    
    ; Check for autostart flag
    If $CmdLine[1] = "--autostart" Then
        ConsoleWrite("Smart Tabs autostart requested via launcher" & @CRLF)
        Local $launchSuccess = LaunchSmartTabsForAutostart()
        
        If Not $launchSuccess Then
            ConsoleWrite("Failed to launch Smart Tabs for autostart" & @CRLF)
            Exit(1)
        EndIf
        
        ConsoleWrite("Smart Tabs autostart completed successfully" & @CRLF)
        Exit(0)
    EndIf
    
    ; Handle context menu request
    If $CmdLine[0] < 2 Or $CmdLine[1] <> "--add-shortcut" Then
        ConsoleWrite("Error: Invalid command. Use --add-shortcut <file-path> or --autostart" & @CRLF)
        Exit(1)
    EndIf
    
    Local $filePath = $CmdLine[2]
    If Not FileExists($filePath) Then
        ConsoleWrite("Error: File does not exist: " & $filePath & @CRLF)
        Exit(1)
    EndIf
    
    ; Try to communicate with existing Smart Tabs instance
    Local $success = SendContextMenuRequest($filePath)
    
    If Not $success Then
        ; If communication failed, try to launch Smart Tabs
        ConsoleWrite("No running Smart Tabs instance found, launching..." & @CRLF)
        Local $launchSuccess = LaunchSmartTabs($filePath)
        
        If Not $launchSuccess Then
            ConsoleWrite("Failed to launch Smart Tabs" & @CRLF)
            Exit(1)
        EndIf
    EndIf
    
    ConsoleWrite("Context menu request processed successfully" & @CRLF)
    Exit(0)
EndFunc

; Send context menu request to running Smart Tabs instance
Func SendContextMenuRequest($filePath)
    ; Try multiple communication methods
    
    ; Method 1: Named pipe (preferred)
    Local $success = SendViaPipe($filePath)
    If $success Then Return True
    
    ; Method 2: Temporary file communication
    $success = SendViaFile($filePath)
    If $success Then Return True
    
    ; Method 3: Registry communication (fallback)
    $success = SendViaRegistry($filePath)
    If $success Then Return True
    
    Return False
EndFunc

; Method 1: Named pipe communication
Func SendViaPipe($filePath)
    Local $pipeName = "\\.\pipe\SmartTabsContextMenu"
    Local $timeout = 1000 ; 1 second timeout

    ; Create request data
    Local $timestamp = @YEAR & @MON & @MDAY & @HOUR & @MIN & @SEC & @MSEC
    Local $requestData = '{"type":"context-menu","filePath":"' & JsonEscape($filePath) & '","timestamp":"' & $timestamp & '"}'

    ; Try to connect to named pipe
    Local $hPipe = _WinAPI_CreateFile($pipeName, 3, 6, 0, 3, 0x80, 0) ; Open existing pipe

    If $hPipe = -1 Then
        ConsoleWrite("Named pipe not available, trying other methods..." & @CRLF)
        Return False
    EndIf

    ; Send request
    Local $bytesWritten = 0
    Local $success = _WinAPI_WriteFile($hPipe, $requestData, StringLen($requestData), $bytesWritten)

    If Not $success Or $bytesWritten = 0 Then
        _WinAPI_CloseHandle($hPipe)
        ConsoleWrite("Failed to write to named pipe" & @CRLF)
        Return False
    EndIf

    ; Read response
    Local $response = ""
    Local $buffer = DllStructCreate("char[4096]")
    Local $bytesRead = 0

    $success = _WinAPI_ReadFile($hPipe, DllStructGetPtr($buffer), 4096, $bytesRead)

    If $success And $bytesRead > 0 Then
        $response = StringLeft(DllStructGetData($buffer, 1), $bytesRead)
    EndIf

    _WinAPI_CloseHandle($hPipe)

    ; Check response
    If StringInStr($response, '"success":true') > 0 Then
        ConsoleWrite("Named pipe communication successful" & @CRLF)
        Return True
    Else
        ConsoleWrite("Named pipe communication failed or invalid response: " & $response & @CRLF)
        Return False
    EndIf
EndFunc

; Method 2: File-based communication
Func SendViaFile($filePath)
    Local $tempDir = @TempDir
    Local $requestFile = $tempDir & "\smarttabs-context-request.txt"
    Local $responseFile = $tempDir & "\smarttabs-context-response.txt"
    
    ; Clean up any existing files
    FileDelete($requestFile)
    FileDelete($responseFile)
    
    ; Create request with current timestamp for uniqueness
    Local $timestamp = @YEAR & @MON & @MDAY & @HOUR & @MIN & @SEC & @MSEC
    Local $requestData = '{"type":"context-menu","filePath":"' & JsonEscape($filePath) & '","timestamp":"' & $timestamp & '"}'
    
    ; Write request file
    Local $result = FileWrite($requestFile, $requestData)
    If Not $result Then Return False
    
    ; Wait for response file (indicates Smart Tabs processed the request)
    Local $waitStart = TimerInit()
    While TimerDiff($waitStart) < $MAX_WAIT_TIME
        If FileExists($responseFile) Then
            Local $response = FileRead($responseFile)
            FileDelete($requestFile)
            FileDelete($responseFile)
            
            ; Check if response indicates success
            If StringInStr($response, '"success":true') > 0 Then
                Return True
            EndIf
            Return False
        EndIf
        Sleep(100)
    WEnd
    
    ; Cleanup on timeout
    FileDelete($requestFile)
    Return False
EndFunc

; Method 3: Registry communication (Windows-specific fallback)
Func SendViaRegistry($filePath)
    Local $regKey = "HKEY_CURRENT_USER\Software\SmartTabs\ContextMenu"
    Local $timestamp = @YEAR & @MON & @MDAY & @HOUR & @MIN & @SEC & @MSEC
    
    ; Write request to registry
    RegWrite($regKey, "Request", "REG_SZ", $filePath)
    RegWrite($regKey, "Timestamp", "REG_SZ", $timestamp)
    RegWrite($regKey, "Status", "REG_SZ", "pending")
    
    ; Wait for status change
    Local $waitStart = TimerInit()
    While TimerDiff($waitStart) < $MAX_WAIT_TIME
        Local $status = RegRead($regKey, "Status")
        If $status = "processed" Then
            ; Clean up
            RegDelete($regKey, "Request")
            RegDelete($regKey, "Timestamp")
            RegDelete($regKey, "Status")
            Return True
        ElseIf $status = "error" Then
            ; Clean up
            RegDelete($regKey, "Request")
            RegDelete($regKey, "Timestamp")
            RegDelete($regKey, "Status")
            Return False
        EndIf
        Sleep(100)
    WEnd
    
    ; Cleanup on timeout
    RegDelete($regKey, "Request")
    RegDelete($regKey, "Timestamp")
    RegDelete($regKey, "Status")
    Return False
EndFunc

; Launch Smart Tabs for autostart (non-elevated)
Func LaunchSmartTabsForAutostart()
    ; Find Smart Tabs executable
    Local $exePath = FindSmartTabsExecutable()
    If Not $exePath Then
        ConsoleWrite("Smart Tabs executable not found" & @CRLF)
        Return False
    EndIf
    
    ; Check if Smart Tabs is already running
    Local $processList = ProcessList($SMART_TABS_EXE_NAME)
    If $processList[0][0] > 0 Then
        ConsoleWrite("Smart Tabs is already running, skipping autostart" & @CRLF)
        Return True
    EndIf
    
    ; Launch Smart Tabs with --autostart flag (will trigger UAC but only once)
    ConsoleWrite("Launching Smart Tabs for autostart..." & @CRLF)
    Local $pid = Run('"' & $exePath & '" --autostart', "", @SW_HIDE)
    
    If $pid > 0 Then
        ConsoleWrite("Smart Tabs launched successfully with PID: " & $pid & @CRLF)
        Return True
    Else
        ConsoleWrite("Failed to launch Smart Tabs (PID: " & $pid & ")" & @CRLF)
        Return False
    EndIf
EndFunc

; Launch Smart Tabs with context menu request
Func LaunchSmartTabs($filePath)
    ; Find Smart Tabs executable
    Local $exePath = FindSmartTabsExecutable()
    If Not $exePath Then
        ConsoleWrite("Smart Tabs executable not found" & @CRLF)
        Return False
    EndIf
    
    ; Launch with original arguments (will trigger UAC but only once)
    Local $pid = Run('"' & $exePath & '" --add-shortcut "' & $filePath & '"', "", @SW_HIDE)
    
    Return $pid > 0
EndFunc

; Find Smart Tabs executable
Func FindSmartTabsExecutable()
    ; Method 1: Same directory as launcher
    Local $sameDirPath = @ScriptDir & "\" & $SMART_TABS_EXE_NAME
    If FileExists($sameDirPath) Then Return $sameDirPath
    
    ; Method 2: Parent directory (if launcher is in subdirectory)
    Local $parentDirPath = StringLeft(@ScriptDir, StringInStr(@ScriptDir, "\", 0, -1) - 1) & "\" & $SMART_TABS_EXE_NAME
    If FileExists($parentDirPath) Then Return $parentDirPath
    
    ; Method 3: Common installation paths
    Local $programFiles = @ProgramFilesDir & "\Smart Tabs\" & $SMART_TABS_EXE_NAME
    If FileExists($programFiles) Then Return $programFiles
    
    Local $programFilesX86 = @ProgramFilesDir & " (x86)\Smart Tabs\" & $SMART_TABS_EXE_NAME
    If FileExists($programFilesX86) Then Return $programFilesX86
    
    ; Method 4: Registry lookup
    Local $regPath = RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Smart Tabs", "InstallPath")
    If $regPath Then
        Local $regExePath = $regPath & "\" & $SMART_TABS_EXE_NAME
        If FileExists($regExePath) Then Return $regExePath
    EndIf
    
    Return ""
EndFunc

; JSON escape helper
Func JsonEscape($text)
    Local $escaped = $text
    $escaped = StringReplace($escaped, "\", "\\")
    $escaped = StringReplace($escaped, '"', '\"')
    $escaped = StringReplace($escaped, @CR, "\r")
    $escaped = StringReplace($escaped, @LF, "\n")
    $escaped = StringReplace($escaped, @TAB, "\t")
    Return $escaped
EndFunc

; WinAPI functions for named pipe communication
Func _WinAPI_CreateFile($sFileName, $iDesiredAccess, $iShareMode, $iSecurityAttributes, $iCreationDisposition, $iFlagsAndAttributes, $hTemplateFile)
    Local $aResult = DllCall("kernel32.dll", "handle", "CreateFileW", "wstr", $sFileName, "dword", $iDesiredAccess, "dword", $iShareMode, "ptr", $iSecurityAttributes, "dword", $iCreationDisposition, "dword", $iFlagsAndAttributes, "handle", $hTemplateFile)
    If @error Then Return SetError(@error, @extended, -1)
    Return $aResult[0]
EndFunc

Func _WinAPI_WriteFile($hFile, $pBuffer, $iToWrite, ByRef $iWritten)
    Local $aResult = DllCall("kernel32.dll", "bool", "WriteFile", "handle", $hFile, "ptr", $pBuffer, "dword", $iToWrite, "dword*", $iWritten, "ptr", 0)
    If @error Then Return SetError(@error, @extended, False)
    Return $aResult[0]
EndFunc

Func _WinAPI_ReadFile($hFile, $pBuffer, $iToRead, ByRef $iRead)
    Local $aResult = DllCall("kernel32.dll", "bool", "ReadFile", "handle", $hFile, "ptr", $pBuffer, "dword", $iToRead, "dword*", $iRead, "ptr", 0)
    If @error Then Return SetError(@error, @extended, False)
    Return $aResult[0]
EndFunc

Func _WinAPI_CloseHandle($hObject)
    Local $aResult = DllCall("kernel32.dll", "bool", "CloseHandle", "handle", $hObject)
    If @error Then Return SetError(@error, @extended, False)
    Return $aResult[0]
EndFunc


