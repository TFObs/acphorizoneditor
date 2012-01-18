VERSION 5.00
Begin VB.Form frmDonate 
   BackColor       =   &H80000008&
   Caption         =   "Make a Donation"
   ClientHeight    =   1800
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4845
   FillStyle       =   0  'Ausgefüllt
   LinkTopic       =   "frmDonate"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   4845
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton cmdDonate 
      BackColor       =   &H0080C0FF&
      Caption         =   "Continue..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      MaskColor       =   &H0080FFFF&
      Style           =   1  'Grafisch
      TabIndex        =   1
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1575
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4335
   End
End
Attribute VB_Name = "frmDonate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Function needed for opening the PayPal-Site
Private Declare Function ShellExecute Lib "shell32.dll" _
        Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal _
        lpOperation As String, ByVal lpFile As String, ByVal _
        lpParameters As String, ByVal lpDirectory As String, _
        ByVal nShowCmd As Long) As Long
        


Private Sub cmdDonate_Click()
Dim result&

result = ShellExecute(Me.hwnd, "Open", _
             "donate.html", "", App.Path, 1)
End Sub

Private Sub cmdDonate_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    cmdDonate.ToolTipText = "click Me!"
End Sub

Private Sub Form_Load()
 Label1.Caption = "If you like the Program, feel free to give " & vbCrLf & _
 "a donation in recognition of all" & vbCrLf & _
 "the work I have put into this Project." & vbCrLf & _
 "Tank you!."
 
 FloatWindow frmDonate.hwnd, True
End Sub

Private Sub Form_Unload(Cancel As Integer)
FloatWindow frmDonate.hwnd, False
If frmMain.mnuchkFloat.Checked Then FloatWindow frmMain.hwnd, True
End Sub


