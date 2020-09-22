VERSION 5.00
Begin VB.Form frmReadOutputExample 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Read Output Example"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10605
   Icon            =   "frmReadOutputExample.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   10605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin prjReadOutput.ReadOutput ReadOutput1 
      Left            =   9240
      Top             =   3000
      _ExtentX        =   1720
      _ExtentY        =   1296
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "E&xecute"
      Height          =   375
      Left            =   9120
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox txtCommand 
      Height          =   285
      Left            =   4200
      TabIndex        =   1
      Text            =   "ping www.google.com"
      Top             =   120
      Width           =   4815
   End
   Begin VB.TextBox txtOutput 
      Height          =   3135
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   600
      Width           =   10335
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "Command to get output from:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3915
   End
End
Attribute VB_Name = "frmReadOutputExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'You may use this code in your project as long as you dont claim its yours ;)

Option Explicit

Private Sub cmdExecute_Click()
    ReadOutput1.SetCommand = txtCommand.Text    'set the command to execute
    ReadOutput1.ProcessCommand                  'launch the command
End Sub

Private Sub Form_Load()

End Sub

Private Sub ReadOutput1_Complete()
    MsgBox "Complete reading output!", vbOKOnly, "Success!" 'command is done
End Sub

Private Sub ReadOutput1_Error(ByVal Error As String, LastDLLError As Long)
    MsgBox "Error!" & vbNewLine & _
            "Description: " & Error & vbNewLine & _
            "LastDLLError: " & LastDLLError, vbCritical, "Error"
End Sub

Private Sub ReadOutput1_GotChunk(ByVal sChunk As String, ByVal LastChunk As Boolean)
    'your probly wondering why I put LastChunk when I already put the Complete event..
    'if you test you'll see that you get chunk by chunk (256 chars), not line by line
    'so if you want to parse those, you'll need to know when it finishes so you can
    'release your last line since you cannot check if its complete by using the event.
    'LastChunk is false if there is more chunks, true if that is the last chunk.
    txtOutput.Text = txtOutput.Text & Replace(Replace(sChunk, Chr(13), ""), Chr(10), vbNewLine)
    'we replace for c/cpp programs because they dont use \c\n they simply use \n so this will support both
    'types of applications
End Sub

Private Sub ReadOutput1_Starting()
    txtOutput.Text = "" 'reset because we dont want to have the old commands output
End Sub
