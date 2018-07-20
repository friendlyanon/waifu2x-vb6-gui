Attribute VB_Name = "U"
Attribute VB_Description = "Author: Dipak Auddy. Email: dauddy@gmail.com"
Option Explicit

Private Const STARTF_USESHOWWINDOW     As Long = &H1
Private Const STARTF_USESTDHANDLES     As Long = &H100
Private Const SW_HIDE                  As Integer = 0
Private Const SW_SHOW                  As Integer = 1
Private Const INFINITE                 As Long = -1&

Private Type SECURITY_ATTRIBUTES
    nLength                                As Long
    lpSecurityDescriptor                   As Long
    bInheritHandle                         As Long
End Type
Private Type STARTUPINFO
    cb                                     As Long
    lpReserved                             As String
    lpDesktop                              As String
    lpTitle                                As String
    dwX                                    As Long
    dwY                                    As Long
    dwXSize                                As Long
    dwYSize                                As Long
    dwXCountChars                          As Long
    dwYCountChars                          As Long
    dwFillAttribute                        As Long
    dwFlags                                As Long
    wShowWindow                            As Integer
    cbReserved2                            As Integer
    lpReserved2                            As Long
    hStdInput                              As Long
    hStdOutput                             As Long
    hStdError                              As Long
End Type
Private Type PROCESS_INFORMATION
    hProcess                               As Long
    hThread                                As Long
    dwProcessId                            As Long
    dwThreadId                             As Long
End Type

Private Declare Function CreatePipe Lib "kernel32" (phReadPipe As Long, _
                                                    phWritePipe As Long, _
                                                    lpPipeAttributes As Any, _
                                                    ByVal nSize As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, _
                                                  lpBuffer As Any, _
                                                  ByVal nNumberOfBytesToRead As Long, _
                                                  lpNumberOfBytesRead As Long, _
                                                  lpOverlapped As Any) As Long
Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, _
                                                                              ByVal lpCommandLine As String, _
                                                                              lpProcessAttributes As Any, _
                                                                              lpThreadAttributes As Any, _
                                                                              ByVal bInheritHandles As Long, _
                                                                              ByVal dwCreationFlags As Long, _
                                                                              lpEnvironment As Any, _
                                                                              ByVal lpCurrentDriectory As String, _
                                                                              lpStartupInfo As STARTUPINFO, _
                                                                              lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, _
                                                             ByVal dwMilliseconds As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, _
                                                            lpExitCode As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public Sub ExecAndCapture(ByVal sCommandLine As String, _
                          cTextBox As TextBox, _
                          Optional ByVal sStartInFolder As String = vbNullString, _
                          Optional bHideWindow As Boolean = True)

    Const BUFSIZE         As Long = 10240
    Dim hPipeRead         As Long
    Dim hPipeWrite        As Long
    Dim sa                As SECURITY_ATTRIBUTES
    Dim si                As STARTUPINFO
    Dim pi                As PROCESS_INFORMATION
    Dim baOutput(BUFSIZE) As Byte
    Dim sOutput           As String
    Dim lBytesRead        As Long
    
    With sa
        .nLength = Len(sa)
        .bInheritHandle = 1
    End With
    
    If CreatePipe(hPipeRead, hPipeWrite, sa, 0) = 0 Then
        Exit Sub
    End If

    With si
        .cb = Len(si)
        .dwFlags = STARTF_USESHOWWINDOW Or STARTF_USESTDHANDLES
        .wShowWindow = IIf(bHideWindow, SW_HIDE, SW_SHOW)
        .hStdOutput = hPipeWrite
        .hStdError = hPipeWrite
    End With
    
    If CreateProcess(vbNullString, sCommandLine, ByVal 0&, ByVal 0&, 1, 0&, ByVal 0&, sStartInFolder, si, pi) Then
        Call CloseHandle(hPipeWrite)
        Call CloseHandle(pi.hThread)
        hPipeWrite = 0
        Do
            DoEvents
            If ReadFile(hPipeRead, baOutput(0), BUFSIZE, lBytesRead, ByVal 0&) = 0 Then
                Exit Do
            End If
            sOutput = Left$(StrConv(baOutput(), vbUnicode), lBytesRead)
            cTextBox.SelText = sOutput
        Loop
        Call CloseHandle(pi.hProcess)
    End If
    Call CloseHandle(hPipeRead)
    Call CloseHandle(hPipeWrite)

End Sub

Public Sub Exec(ByVal sCommandLine As String, _
                          Optional ByVal sStartInFolder As String = vbNullString, _
                          Optional bHideWindow As Boolean = True)

    Dim si  As STARTUPINFO
    Dim pi  As PROCESS_INFORMATION
    Dim ret As Long

    With si
        .cb = Len(si)
        .wShowWindow = IIf(bHideWindow, SW_HIDE, SW_SHOW)
    End With
    
    ret = CreateProcess(vbNullString, sCommandLine, ByVal 0&, ByVal 0&, 1, 0&, ByVal 0&, sStartInFolder, si, pi)
    ret = WaitForSingleObject(pi.hProcess, INFINITE)
    Call GetExitCodeProcess(pi.hProcess, ret)
    Call CloseHandle(pi.hThread)
    Call CloseHandle(pi.hProcess)

End Sub

