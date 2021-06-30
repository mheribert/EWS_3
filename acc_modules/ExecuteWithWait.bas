Option Compare Database
Option Explicit

#If Win64 Then
'    Private Declare PtrSafe Function WaitForSingleObject Lib "Kernel32" (ByVal hHandle As LongPtr, ByVal dwMilliseconds As LongLong) As LongLong
'    Private Declare PtrSafe Function CloseHandle Lib "Kernel32" (ByVal hObject As LongPtr) As LongLong
'    Private Declare PtrSafe Function CreateProcess Lib "Kernel32" (ByVal lpAppName As LongPtr, ByVal lpCmdLine As LongPtr, ByVal lpProcAttr As LongLong, ByVal lpThreadAttr As LongLong, ByVal lpInheritedHandle As LongPtr, ByVal lpCreationFlags As LongLong, ByVal lpEnv As LongLong, ByVal lpCurDir As LongPtr, ByVal lpStartupInfo As LongPtr, ByVal lpProcessInfo As LongPtr) As LongLong
'
'    Private Type STARTUPINFO
'           cb As LongLong
'           lpReserved As String
'           lpDesktop As String
'           lpTitle As String
'           dwX As LongLong
'           dwY As LongLong
'           dwXSize As LongLong
'           dwYSize As LongLong
'           dwXCountChars As LongLong
'           dwYCountChars As LongLong
'           dwFillAttribute As LongLong
'           dwFlags As LongLong
'           wShowWindow As Long
'           cbReserved2 As Long
'           lpReserved2 As Long
'           hStdInput As LongPtr
'           hStdOutput As LongPtr
'           hStdError As LongPtr
'    End Type
'
'    Private Type PROCESS_INFORMATION
'          hProcess As LongPtr
'          hThread As LongPtr
'          dwProcessID As LongLong
'          dwThreadID As LongLong
'    End Type
'
'    Private Const NORMAL_PRIORITY_CLASS  As LongLong = &H20&
'    Private Const INFINITE As LongLong = -1&
'    Private Const WAIT_TIMEOUT As LongLong = 258&
#Else

    Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
    Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
    Private Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long

    Private Type STARTUPINFO
        cb As Long
        lpReserved As String
        lpDesktop As String
        lpTitle As String
        dwX As Long
        dwY As Long
        dwXSize As Long
        dwYSize As Long
        dwXCountChars As Long
        dwYCountChars As Long
        dwFillAttribute As Long
        dwFlags As Long
        wShowWindow As Integer
        cbReserved2 As Integer
        lpReserved2 As Long
        hStdInput As Long
        hStdOutput As Long
        hStdError As Long
    End Type

    Private Type PROCESS_INFORMATION
        hProcess As Long
        hThread As Long
        dwProcessID As Long
        dwThreadID As Long
    End Type
   
    Private Const NORMAL_PRIORITY_CLASS = &H20&
    Private Const INFINITE = -1&
    Private Const WAIT_TIMEOUT As Long = 258&
#End If

   
Public Sub ExecCmd(cmdline$)
    Dim proc As PROCESS_INFORMATION
    Dim Start As STARTUPINFO
    Dim ReturnValue As Integer
    ' Initialisiert die STARTUPINFO Struktur:
    Start.cb = Len(Start)
    ' Startet die Shell-Anwendung:
    ReturnValue = CreateProcessA(0&, cmdline$, 0&, 0&, 1&, NORMAL_PRIORITY_CLASS, 0&, 0&, Start, proc)
    ' Wartet bis Shell-Anwendung geschlossen ist:
    Do
        ReturnValue = WaitForSingleObject(proc.hProcess, 0)
        DoEvents
    Loop Until ReturnValue <> 258
    ReturnValue = CloseHandle(proc.hProcess)
End Sub


Sub helper()
    Dim db As Database
    Dim re As Recordset
    Dim quell As String
    Set db = CurrentDb
    Set re = db.OpenRecordset("SELECT * FROM WR_Einteilung ORDER BY WR_Anzeige;")
    re.MoveFirst
    quell = ""
    Do Until re.EOF
        quell = quell & re!WR_Anzahl & ";" & """" & re!WR_Anzeige & """;"
    
    
    
    
        re.MoveNext
    Loop
    Debug.Print quell
End Sub
