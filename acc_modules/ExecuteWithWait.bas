Option Compare Database
Option Explicit

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
   
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

Private Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Const NORMAL_PRIORITY_CLASS = &H20&

Private Const INFINITE = -1&

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
