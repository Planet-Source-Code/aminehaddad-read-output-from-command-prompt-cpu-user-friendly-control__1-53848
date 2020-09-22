VERSION 5.00
Begin VB.UserControl ReadOutput 
   BackColor       =   &H80000007&
   ClientHeight    =   765
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1020
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   765
   ScaleWidth      =   1020
   Begin VB.Image imgIcon 
      Height          =   735
      Left            =   0
      Picture         =   "ReadOutput.ctx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   975
   End
End
Attribute VB_Name = "ReadOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'You may use this code in your project as long as you dont claim its yours ;)

'This program reads the output of CLI (Command Line Interface) Applications.
'Examples of CLI Applications are:
'   -PING.EXE
'   -NETSTAT
'   -TRACERT
'This program will grab the output and call events so that you can process the commands.
'NOTE:  I got about 50% of this code from some site about 2 years ago, just found it and fixed some bugs
'       and transformed it into a user-friendly control.
'
'Please vote if you like, complaint about bugs if you find any, but I also appreciate comments ;)
'Thanks again
'-Endra

Option Explicit 'force declarations of variables

'private variables
Dim sCommand As String
Dim bProcessing As Boolean

'Public Events
Public Event Error(ByVal Error As String, LastDLLError As Long) 'Errors go here
Public Event GotChunk(ByVal sChunk As String, ByVal LastChunk As Boolean)                   'Chunk Output detected, launch this event
Public Event Complete()                                         'Raise event when its done reading output
Public Event Starting()                                         'Raised when you can start the reading

'The following are all API Calls and Types.
'You could probably find more information on them if you google them so I wont comment them at all.
Private Declare Function CreatePipe Lib "kernel32" ( _
    phReadPipe As Long, _
    phWritePipe As Long, _
    lpPipeAttributes As Any, _
    ByVal nSize As Long) As Long
      
Private Declare Function ReadFile Lib "kernel32" ( _
    ByVal hFile As Long, _
    ByVal lpBuffer As String, _
    ByVal nNumberOfBytesToRead As Long, _
    lpNumberOfBytesRead As Long, _
    ByVal lpOverlapped As Any) As Long
      
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Type STARTUPINFO
    cb As Long
    lpReserved As Long
    lpDesktop As Long
    lpTitle As Long
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
    dwProcessId As Long
    dwThreadID As Long
End Type

Private Declare Function CreateProcessA Lib "kernel32" (ByVal _
    lpApplicationName As Long, ByVal lpCommandLine As String, _
    lpProcessAttributes As Any, lpThreadAttributes As Any, _
    ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
    ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, _
    lpStartupInfo As Any, lpProcessInformation As Any) As Long
      
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

'The following are simply constants that dont need changing during the program.
'DO NOT EDIT THESE!

Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const STARTF_USESTDHANDLES = &H100&
Private Const STARTF_USESHOWWINDOW = &H1
Private Const SW_HIDE = 0

Private Sub UserControl_Initialize()
    On Error Resume Next
    'doesnt error out of stack space
    UserControl.Height = imgIcon.Height
    UserControl.Width = imgIcon.Width
    bProcessing = False
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    'doesnt error out of stack space
    UserControl.Height = imgIcon.Height
    UserControl.Width = imgIcon.Width
End Sub

'The following function executes the command line and returns the output via events
Private Function ExecuteApp(sCmdline As String) As String
    Dim proc As PROCESS_INFORMATION, ret As Long
    Dim start As STARTUPINFO
    Dim sa As SECURITY_ATTRIBUTES
    Dim hReadPipe As Long 'The handle used to read from the pipe.
    Dim hWritePipe As Long 'The pipe where StdOutput and StdErr will be redirected to.
    Dim sOutput As String
    Dim lngBytesRead As Long, sBuffer As String * 256
    bProcessing = True
    sa.nLength = Len(sa)
    sa.bInheritHandle = True
      
    ret = CreatePipe(hReadPipe, hWritePipe, sa, 0)
    If ret = 0 Then
        bProcessing = False
        RaiseEvent Error("CreatePipe failed.", Err.LastDLLError)
        Exit Function
    End If

    start.cb = Len(start)
    start.dwFlags = STARTF_USESTDHANDLES Or STARTF_USESHOWWINDOW
    ' Redirect the standard output and standard error to the same pipe
    start.hStdOutput = hWritePipe
    start.hStdError = hWritePipe
    start.wShowWindow = SW_HIDE
       
    ' Start the shelled application:
    ' if you program has to work only on NT you don't need the "conspawn "
    'ret = CreateProcessA(0&, "conspawn " & sCmdline, sa, sa, True, NORMAL_PRIORITY_CLASS, 0&, 0&, start, proc)
    ret = CreateProcessA(0&, Environ("ComSpec") & " /c " & sCmdline, sa, sa, True, NORMAL_PRIORITY_CLASS, 0&, 0&, start, proc)
    If ret = 0 Then
        bProcessing = False
        RaiseEvent Error("CreateProcess failed.", Err.LastDLLError)
        Exit Function
    End If
   
    ' The handle wWritePipe has been inherited by the shelled application
    ' so we can close it now
    CloseHandle hWritePipe

    ' Read the characters that the shelled application
    ' has outputed 256 characters at a time
    RaiseEvent Starting
    Do
        DoEvents
        ret = ReadFile(hReadPipe, sBuffer, 256, lngBytesRead, 0&)
        sOutput = Left$(sBuffer, lngBytesRead)
        If ret = 0 Then
            RaiseEvent GotChunk(sOutput, True)  'no more chunks to read
            RaiseEvent Complete
            Exit Do
        Else
            RaiseEvent GotChunk(sOutput, False) 'more chunks to read.
        End If
    Loop While ret <> 0 ' if ret = 0 then there is no more characters to read
    
    CloseHandle proc.hProcess
    CloseHandle proc.hThread
    CloseHandle hReadPipe
    bProcessing = False
End Function

Public Property Let SetCommand(ByVal sCommandVal As String)
    sCommand = sCommandVal
End Property

Public Property Get SetCommand() As String
    SetCommand = sCommand
End Property

Public Sub ProcessCommand()
    If Len(sCommand) = 0 Then
        RaiseEvent Error("Invalid Command.", 1200)
        Exit Sub
    End If
    If bProcessing = True Then
        RaiseEvent Error("Currently processing a command!", 1201)
        Exit Sub
    End If
    ExecuteApp """" & sCommand & """"
End Sub

