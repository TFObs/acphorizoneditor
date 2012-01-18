VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'Kein
   Caption         =   "ACP Horizon Editor     (c) J.Hanisch 2012"
   ClientHeight    =   7470
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   9135
   FillColor       =   &H80000007&
   FillStyle       =   0  'Ausgefüllt
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   9135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin MSFlexGridLib.MSFlexGrid GridMark 
      Height          =   6015
      Left            =   8640
      TabIndex        =   39
      Top             =   720
      Width           =   225
      _ExtentX        =   397
      _ExtentY        =   10610
      _Version        =   393216
      Cols            =   1
      BackColor       =   -2147483641
      BackColorFixed  =   -2147483641
      BackColorSel    =   -2147483641
      ForeColorSel    =   -2147483641
      BackColorBkg    =   0
      GridColor       =   0
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   0
      GridLinesFixed  =   3
      ScrollBars      =   2
      BorderStyle     =   0
   End
   Begin VB.CommandButton cmdInterpolate 
      BackColor       =   &H00000000&
      Height          =   975
      Left            =   2520
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmMain.frx":08CA
      Style           =   1  'Grafisch
      TabIndex        =   35
      ToolTipText     =   "Interpolate Values"
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdPause 
      BackColor       =   &H0099A8AC&
      Enabled         =   0   'False
      Height          =   975
      Left            =   1560
      MaskColor       =   &H00000000&
      Picture         =   "frmMain.frx":1194
      Style           =   1  'Grafisch
      TabIndex        =   34
      ToolTipText     =   "Pause the Recording"
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdContinue 
      BackColor       =   &H80000007&
      Height          =   975
      Left            =   1560
      MaskColor       =   &H00000000&
      Picture         =   "frmMain.frx":1A5E
      Style           =   1  'Grafisch
      TabIndex        =   33
      ToolTipText     =   "Continue Recording Horizon"
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdConnect 
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
      Height          =   975
      Left            =   240
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmMain.frx":2328
      Style           =   1  'Grafisch
      TabIndex        =   17
      ToolTipText     =   "Connect to the Scope via ACP"
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000007&
      Caption         =   "open..."
      ForeColor       =   &H8000000E&
      Height          =   1095
      Left            =   4200
      TabIndex        =   14
      Top             =   0
      Width           =   2295
      Begin MSComDlg.CommonDialog diagOpenSave 
         Left            =   1800
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H80000007&
         Caption         =   "TheSky Horizon-file"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   37
         Top             =   740
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H80000007&
         Caption         =   "from File"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H80000007&
         Caption         =   "current ACP-Horizon"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   490
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000007&
      Caption         =   "save..."
      ForeColor       =   &H8000000E&
      Height          =   1095
      Left            =   4200
      TabIndex        =   11
      Top             =   1080
      Width           =   2295
      Begin VB.OptionButton Option1 
         BackColor       =   &H80000007&
         Caption         =   "as TheSky Horizon-File"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   38
         Top             =   720
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H80000007&
         Caption         =   "Save and use with ACP"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H80000007&
         Caption         =   "to File"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdOpen 
      Appearance      =   0  '2D
      BackColor       =   &H00000000&
      Height          =   735
      Left            =   3480
      MaskColor       =   &H00000000&
      Picture         =   "frmMain.frx":2996
      Style           =   1  'Grafisch
      TabIndex        =   10
      ToolTipText     =   "Open Horizon from File ore current ACP-Hor."
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00000000&
      Height          =   735
      Left            =   3480
      MaskColor       =   &H00000000&
      Picture         =   "frmMain.frx":3260
      Style           =   1  'Grafisch
      TabIndex        =   9
      ToolTipText     =   "Save the Horizon to File or apply directly to ACP"
      Top             =   1200
      Width           =   615
   End
   Begin VB.PictureBox picHorizon 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   480
      ScaleHeight     =   4275
      ScaleWidth      =   5955
      TabIndex        =   8
      Top             =   2280
      Width           =   6015
      Begin VB.Line LineAlt 
         BorderColor     =   &H80000003&
         Index           =   7
         X1              =   0
         X2              =   6000
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Line LineAlt 
         BorderColor     =   &H80000003&
         Index           =   6
         X1              =   0
         X2              =   6000
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line LineAlt 
         BorderColor     =   &H80000003&
         Index           =   5
         X1              =   0
         X2              =   6000
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Line LineAlt 
         BorderColor     =   &H80000003&
         Index           =   4
         X1              =   0
         X2              =   6000
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Line LineAlt 
         BorderColor     =   &H80000003&
         Index           =   3
         X1              =   0
         X2              =   6000
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Line LineAlt 
         BorderColor     =   &H80000003&
         Index           =   2
         X1              =   0
         X2              =   6000
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Line LineAlt 
         BorderColor     =   &H80000003&
         Index           =   1
         X1              =   -120
         X2              =   5880
         Y1              =   3120
         Y2              =   3120
      End
      Begin VB.Line LineAlt 
         BorderColor     =   &H80000003&
         Index           =   0
         X1              =   0
         X2              =   6000
         Y1              =   3720
         Y2              =   3720
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000FF&
         BorderStyle     =   4  'Strich-Punkt
         BorderWidth     =   3
         Height          =   195
         Left            =   2640
         Shape           =   1  'Quadrat
         Top             =   2040
         Width           =   195
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         X1              =   2880
         X2              =   2880
         Y1              =   6000
         Y2              =   0
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         X1              =   4440
         X2              =   4440
         Y1              =   4320
         Y2              =   0
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         X1              =   1080
         X2              =   1080
         Y1              =   4320
         Y2              =   0
      End
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H000040C0&
      Caption         =   "Clear Horizon"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6840
      Style           =   1  'Grafisch
      TabIndex        =   7
      ToolTipText     =   "Clear the current Horizon"
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   1455
   End
   Begin VB.TextBox txtInput 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'Kein
      Height          =   285
      Left            =   7800
      TabIndex        =   3
      Top             =   2040
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid GridHoriz 
      Height          =   6255
      Left            =   6720
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   11033
      _Version        =   393216
      BackColor       =   -2147483635
      ForeColor       =   65535
      BackColorFixed  =   33023
      BackColorBkg    =   0
      GridColor       =   8421504
      AllowBigSelection=   0   'False
      FocusRect       =   2
      HighLight       =   2
      ScrollBars      =   2
      AllowUserResizing=   1
      BorderStyle     =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdDisconnect 
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
      Height          =   975
      Left            =   240
      Picture         =   "frmMain.frx":356A
      Style           =   1  'Grafisch
      TabIndex        =   0
      ToolTipText     =   "Disconnect the Scope"
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00000000&
      Caption         =   "5° Check"
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
      Height          =   615
      Left            =   8520
      TabIndex        =   40
      Top             =   240
      Width           =   615
   End
   Begin VB.Label lblStatus 
      BackColor       =   &H80000012&
      Caption         =   "Idle"
      ForeColor       =   &H0000C0C0&
      Height          =   255
      Left            =   120
      TabIndex        =   36
      Top             =   7200
      Width           =   7695
   End
   Begin VB.Label lblAzim 
      BackStyle       =   0  'Transparent
      Caption         =   "N"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Index           =   4
      Left            =   6360
      TabIndex        =   32
      Top             =   6480
      Width           =   375
   End
   Begin VB.Label lblAzim 
      BackStyle       =   0  'Transparent
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Index           =   3
      Left            =   4680
      TabIndex        =   31
      Top             =   6480
      Width           =   375
   End
   Begin VB.Label lblAzim 
      BackStyle       =   0  'Transparent
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Index           =   2
      Left            =   3240
      TabIndex        =   30
      Top             =   6480
      Width           =   375
   End
   Begin VB.Label lblAzim 
      BackStyle       =   0  'Transparent
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Index           =   1
      Left            =   1440
      TabIndex        =   29
      Top             =   6480
      Width           =   375
   End
   Begin VB.Label lblAzim 
      BackStyle       =   0  'Transparent
      Caption         =   "N"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Index           =   0
      Left            =   480
      TabIndex        =   28
      Top             =   6480
      Width           =   375
   End
   Begin VB.Label lblAlt 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   27
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label lblAlt 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   26
      Top             =   2400
      Width           =   255
   End
   Begin VB.Label lblAlt 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   25
      Top             =   5760
      Width           =   255
   End
   Begin VB.Label lblAlt 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   24
      Top             =   5160
      Width           =   255
   End
   Begin VB.Label lblAlt 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   23
      Top             =   4800
      Width           =   255
   End
   Begin VB.Label lblAlt 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   22
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label lblAlt 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   21
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label lblAlt 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   20
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label lblAlt 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   19
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label lblAlt 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   18
      Top             =   6240
      Width           =   255
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000007&
      Caption         =   "Altitude:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      Caption         =   "Horizon:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label lblElev 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00000000&
      Caption         =   "- - -"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label lblHoriz 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00000000&
      Caption         =   "- - -"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Menu mnuExit 
      Caption         =   "EXIT"
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "About"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
   End
   Begin VB.Menu mnuDonate 
      Caption         =   "Donate!"
   End
   Begin VB.Menu empty 
      Caption         =   ""
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu mnuFloat 
      Caption         =   "Float Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuchkFloat 
         Caption         =   "Always on Top"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'Api for the Helpfile
Private Declare Function HtmlHelp Lib "hhctrl.ocx" _
            Alias "HtmlHelpA" (ByVal hwndCaller As Long, _
            ByVal pszFile As String, ByVal uCommand As _
            Long, ByVal dwData As Long) As Long

Private Const HH_DISPLAY_TOPIC = &H0
Private Const HH_CLOSE_ALL As Long = &H12

Dim isPaused As Boolean
Dim ScopeConnectedAtStartup As Boolean
Dim ErrCatcher As DebugHelper



Private Sub cmdClear_Click()
    picHorizon.Cls
        
    For x = 1 To 180
        GridHoriz.TextMatrix(x, 1) = "--"
    Next x
        
      
    lblHoriz.Caption = "- - -"
    lblElev.Caption = "- - -"
        
    CheckHorizon (True)

End Sub

Private Sub cmdContinue_Click()             'Continue Scope-Recording

 isPaused = False
 cmdPause.Visible = True
 cmdContinue.Visible = False
 lblStatus.Caption = "Application Status: Scope-Tracking: " & IIf(isPaused, "Idle/Paused", "in Progress")
  
End Sub

'interpolate when in Inerpolate-mode
Private Sub Interpolate(x1, y1, x2, y2)
Dim x3, y3
Dim counter As Integer
  
  'For counter = (x1 / 2) + 2 To (x2 / 2)
 For counter = x1 To x2 Step 2
    x3 = counter
    y3 = y1 + ((x3 - x1) / (x2 - x1)) * (y2 - y1)
    
    If x3 / 2 + 1 = 181 Then
        GridHoriz.TextMatrix(1, 1) = Format(y3, "0.0")
    Else
        GridHoriz.TextMatrix(x3 / 2 + 1, 1) = Format(y3, "0.0")
    End If
 Next counter
 
  
End Sub

Private Sub cmdinterpolate_click()
Dim dicInterpol As Dictionary
Dim colInterpol As Collection
Dim firstValue As Integer
Dim lastValue As Integer

Set colInterpol = New Collection
Set dicInterpol = New Dictionary


    'If Not IsNumeric(GridHoriz.TextMatrix(1, 1)) Then GridHoriz.TextMatrix(1, 1) = "0.0"
     '     If Not IsNumeric(GridHoriz.TextMatrix(180, 1)) Then GridHoriz.TextMatrix(180, 1) = "0.0"

    For x = 1 To 180
        If IsNumeric(GridHoriz.TextMatrix(x, 1)) Then
            dicInterpol.Add GridHoriz.TextMatrix(x, 0), GridHoriz.TextMatrix(x, 1)
            colInterpol.Add GridHoriz.TextMatrix(x, 0)
            Else
                If GridHoriz.TextMatrix(x, 1) <> "--" Or dicInterpol.Count = 0 Then
                    MsgBox "Error ! Non numerical Value" & vbCrLf & vbCrLf & "please check your Value at " & _
                    GridHoriz.TextMatrix(x, 0) & "°" & vbCrLf & "Cannot Continue...", vbExclamation
                    Exit Sub
                End If
        End If
        
    Next x
    
    firstValue = colInterpol(1): lastValue = colInterpol(colInterpol.Count)
    
 
     
     
        For x = 1 To colInterpol.Count - 1
            Interpolate colInterpol(x), dicInterpol(colInterpol(x)), colInterpol(x + 1), dicInterpol(colInterpol(x + 1))
        Next x
    
   
      If Not (firstValue = 0 And lastValue = 358) Then
        Dim x1, x2, x3, y1, y2, y3
        
        x1 = 0
        y1 = CDbl(dicInterpol(CStr(lastValue)))
        x2 = 358 - lastValue + firstValue
        y2 = CDbl(dicInterpol(CStr(firstValue)))
        x3 = (358 - lastValue) - 2
        y3 = y1 + ((x3 - x1) / (x2 - x1)) * (y2 - y1)
        
        If lastValue <> 358 Then
            Interpolate lastValue, y1, 358, y3
        End If
        
        If firstValue <> 0 Then
            Interpolate 0, y1 + ((x3 + 4 - x1) / (x2 - x1)) * (y2 - y1), firstValue, y2
        End If
        
        
   
   End If
     'now draw the horizon
     DrawHorizon
         
     lblStatus.Caption = "Application Status: Scope-Tracking: " & IIf(isPaused, "Idle/Paused", "in Progress")
  
End Sub



Private Sub cmdOpen_Click()
Dim openFile As String
Dim HorVals

Dim fs As FileSystemObject
Dim openStream As TextStream

    If Option1(2).value = True Then Call getACPHorizon  'Get current Horizon from Registry
    
    If Option1(3).value = True Then                     'Save Horizon to file
    
        With diagOpenSave                               'Open the Dialog
            .Filter = "Horizon-files (*.hor)|*.hor"
            .InitDir = App.Path
            .MaxFileSize = 2000
            .DialogTitle = "load existing Horizon"
            .ShowOpen
        openFile = .FileName
        End With
        
        If Not openFile = "" Then                       'proceed if file was chosen
            Set fs = New FileSystemObject
            Set openStream = fs.OpenTextFile(openFile)
            
                HorVals = openStream.ReadAll
                HorVals = Split(HorVals, " ")
                
                'save values to Grid
                For x = 0 To UBound(HorVals)
                    frmMain.GridHoriz.TextMatrix(x + 1, 1) = HorVals(x)
                Next x
                
            openStream.Close
                
            Set openStream = Nothing
            Set fs = Nothing
            
        End If
     
     End If
     
     If Option1(4).value = True Then                     'Get TheSky Horizon to file
    
        With diagOpenSave                               'Open the Dialog
            .Filter = "TheSky Horizon-files (*.hrz)|*.hrz"
            .InitDir = App.Path
            .MaxFileSize = 2000
            .DialogTitle = "load existing TheSky Horizon"
            .ShowOpen
        openFile = .FileName
        End With
        
        If Not openFile = "" Then Call getTheSkyFile(openFile)  'proceed if file was chosen
            
            
        End If
        
        frmMain.DrawHorizon         'Draw new Horizon
        
    CheckHorizon (True)
   
End Sub

Private Sub cmdPause_Click()        'Pause the Scope-Recording

    isPaused = True
    cmdContinue.Visible = True
    cmdPause.Visible = False
    
    lblStatus.Caption = "Application Status: Scope-Tracking: " & IIf(isPaused, "Idle/Paused", "in Progress")
  
End Sub

Private Sub cmdSave_Click()
Dim saveFile As String
Dim saveFileX As String 'TheSkyX - file
Dim HorVals

If CheckHorizon(False) = False Then Exit Sub

Dim fs As FileSystemObject
Dim saveStream As TextStream

    If Option1(1) = True Then Call saveACPHorizon              'Save Horizon to Registry
 
    If Option1(0).value = True Then                             'Save to file
    
        With diagOpenSave                                       'open the Dialog
            .Filter = "Horizon-files (*.hor)|*.hor"
            .InitDir = App.Path
            .MaxFileSize = 2000
            .DialogTitle = "Save Horizon"
            .ShowOpen
            saveFile = .FileName
        End With
    
        If Not saveFile = "" Then                               'Proceed if file was chosen
            Set fs = New FileSystemObject
            Set saveStream = fs.CreateTextFile(saveFile)
        
                HorVals = ""
                
                'Read values from Grid and store to Horvals
                For x = 1 To 180
                    If IsNumeric(frmMain.GridHoriz.TextMatrix(x, 1)) Then
                    HorVals = HorVals & (Format(frmMain.GridHoriz.TextMatrix(x, 1), "0.0") & " ")
                    Else: HorVals = HorVals & "0.0 "
                    End If
                Next x
                
                saveStream.Write Trim(HorVals)  'save Values to file
                
                saveStream.Close
                
            Set saveStream = Nothing
            Set fs = Nothing
        End If
     
  End If
  
  If Option1(5).value = True Then                             'Save to TheSkyfile
    
        With diagOpenSave                                       'open the Dialog
            .Filter = "TheSky Horizon-files (*.hrz)|*.hrz"
            .InitDir = App.Path
            .MaxFileSize = 2000
            .DialogTitle = "Save TheSky Horizon-file"
            .ShowOpen
            saveFile = .FileName
            saveFileX = Left(.FileName, Len(.FileName) - 4) & "_X.hrz"
        End With
    
        If Not saveFile = "" Then Call saveTheSkyFile(saveFile, saveFileX)                             'Proceed if file was chosen
            
             
  End If
End Sub

'connects to the scope via ACP
Private Sub cmdConnect_Click()
Dim azimuth As Double
Dim altitude As Double

abort = False

On Error GoTo errhandler
scope.Connected = True

cmdPause_Click  'Startup in Paused-Mode

lblStatus.Caption = "Application Status: Scope-Tracking: " & IIf(isPaused, "Idle/Paused", "in Progress")
  Shape1.Visible = True

cmdConnect.Visible = False
cmdDisconnect.Visible = True
cmdPause.Enabled = True
cmdPause.BackColor = QBColor(0)

    While Not abort     'loop until the "Disconnect-Button is clicked

        On Error Resume Next
        DoEvents
        
        'Label1.Caption = "  Scope:"
        
        If Not isPaused Then DrawHorizon  'Don't update if the Paused-Button is clicked
            
            WaitForMilliseconds 500     'show the Scope position even if Paused
            azimuth = scope.azimuth
            altitude = scope.altitude

            Shape1.Left = azimuth - 9.5
            Shape1.Top = 90 - altitude - 9.5
            
            If Not isPaused Then
                Label1.Caption = "Scope:"
                If Round(azimuth / 2, 0) = 180 Then     'Special treatment (360° = 0°)
                    GridHoriz.TextMatrix(1, 1) = Format(altitude, "0.0")
                Else
                    GridHoriz.TextMatrix(Round(azimuth / 2, 0) + 1, 1) = Format(altitude, "0.0")
                End If
            
                If Round((azimuth / 2), 0) - 11 < 1 Then
                    GridHoriz.TopRow = 1
                Else
                    GridHoriz.TopRow = Round((azimuth / 2), 0) - 11
                End If
                
        Else: Label1.Caption = "Horizon:"
        
        End If
            
       'Show Values on the Form
        lblHoriz.Caption = (Int(Round((azimuth / 2), 0)) * 2) & " °"
        lblElev.Caption = Format(altitude, "0.0") & " °"
       
    Wend

    'scope.Connected = False
    'Set scope = Nothing
    
    Shape1.Visible = False
    
    cmdDisconnect.Visible = False
    cmdConnect.Visible = True
    cmdContinue.Visible = False
    cmdPause.Visible = True
    cmdPause.Enabled = False
    cmdPause.BackColor = &H99A8AC
    lblHoriz.Caption = "- - -"
    lblElev.Caption = "- - -"
    
    DrawHorizon
    Exit Sub
errhandler:
FloatWindow frmMain.hwnd, False
Dim result
    If Err.Number = 462 Then
        Err.Clear
        result = MsgBox("Lost Connection to the Scope" & vbCrLf & "Do you want to try to reconnect?", vbYesNo + vbExclamation, "Connection lost...")
        If result = 7 Then
            Exit Sub
        Else
            'Set scope = New acp.Telescope'<<<--DON'T DO This: ACP 6.0 changed ActiveX interfaces: No early binding!!!
            
            Set scope = CreateObject("ACP.Telescope")
            
            cmdConnect_Click
        End If
End If
   If mnuchkFloat.Checked Then FloatWindow frmMain.hwnd, True
End Sub

Private Sub cmddisconnect_Click()
Dim result

If cmdContinue.Visible = True Then
FloatWindow frmMain.hwnd, False
 result = MsgBox("There is a Recording running which is paused" & vbCrLf & "Do you REALLY want to disconnect?", vbYesNo + vbQuestion, "Recording in Progress")
 If result = 7 Then Exit Sub
   If mnuchkFloat.Checked Then FloatWindow frmMain.hwnd, True
End If
 
 abort = True
 lblStatus.Caption = "Application Status: Scope-Tracking: " & IIf(isPaused, "Idle/Paused", "in Progress")
  
End Sub



Private Sub Form_Load()
On Error GoTo errhandler
ErrCatcher.module = "Form_Load"
ErrCatcher.place = "Beginning"

If App.PrevInstance Then Exit Sub

DisableCloseButton Me.hwnd  'Disable the Closed button to enable question to disconnect
   
   ErrCatcher.module = "Form_Load"
   ErrCatcher.place = "Line 10"
   
    isPaused = True
    
    'Set scope = New acp.Telescope'<<<--DON'T DO This: ACP 6.0 changed ActiveX interfaces: No early binding!!!
    Set scope = CreateObject("ACP.Telescope")
    
    If Not scope.Connected Then
        ScopeConnectedAtStartup = False
    Else
        ScopeConnectedAtStartup = True
    End If

ErrCatcher.place = "GridHoriz"
   
With GridHoriz
    .ColWidth(0) = 800
    .ColWidth(1) = 600
    .ColAlignment(1) = flexAlignCenterCenter
    .Width = 1720
    .Rows = 181
    GridMark.Rows = 180
    GridMark.FixedRows = 0
    GridMark.ColAlignment(0) = flexAlignCenterCenter
    GridMark.ColWidth(0) = 300
    GridMark.Col = 0: GridMark.Row = 0: GridMark.CellBackColor = &H8000&
    
    .TextMatrix(0, 0) = "Azim.[°]"
    .TextMatrix(0, 1) = "Alti.[°]"
    
    For x = 1 To 180
        .TextMatrix(x, 0) = (x - 1) * 2
        .TextMatrix(x, 1) = "--"
        If x < 180 Then GridMark.Row = x: GridMark.CellBackColor = &H8000&
        
    Next x
    

End With

ErrCatcher.place = "picHorizon"
With picHorizon
    .DrawWidth = 2
    .ScaleMode = 3
    .ScaleWidth = 360
    .ScaleHeight = 90
    .ForeColor = QBColor(7)

'Labels for Azimuth
For x = 0 To 4
    lblAzim(x).Left = (.Left) - 100 + (.Width / 4) * x '-8
    lblAzim(x).Top = .Top + .Height + 100 '0
Next x

End With

ErrCatcher.place = "Line1"
With Line1
    .x1 = 180
    .x2 = .x1
    .y1 = 0
    .y2 = 90
End With

With Line2
    .x1 = 90
    .x2 = .x1
    .y1 = 0
    .y2 = 90
End With

With Line3
    .x1 = 270
    .x2 = .x1
    .y1 = 0
    .y2 = 90
End With

For x = 0 To 9
    If x <= 7 Then
        With LineAlt(x)
            .x1 = 0
            .x2 = 360
            .y1 = (x + 1) * 10
            .y2 = .y1
        End With
    End If
        
    With lblAlt(x)
        .FontBold = True
        .BackStyle = 0
        .ForeColor = &H8000000F
        .Caption = (x) * 10
        .Left = picHorizon.Left - 300
        .Top = (picHorizon.Top + picHorizon.Height) - ((picHorizon.Height / 9) * x) - 100
    End With
Next x
 
ErrCatcher.place = "Shape"
Shape1.Height = 20
Shape1.Width = 20
Shape1.Visible = False

ErrCatcher.place = "lblStatus"
lblStatus.Caption = "Application Status: Scope-Tracking: " & IIf(isPaused, "Idle/Paused", "in Progress")
If mnuchkFloat.Checked Then FloatWindow Me.hwnd, True
Exit Sub

errhandler:
MsgBox "Error " & Err.Number & " " & Err.Description & vbCrLf & vbCrLf & _
"Procedure: " & ErrCatcher.module & vbCrLf & "Place: " & ErrCatcher.place, vbCritical, "Error Catched"
  End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 End
End Sub

Private Sub GridHoriz_Scroll()
    txtInput.Visible = False
    GridMark.TopRow = GridHoriz.TopRow - 1
End Sub

Private Sub mnuAbout_Click()
FloatWindow frmMain.hwnd, False
    MsgBox "ACP Horizon Editor" & vbCrLf & vbCrLf & "Written by Jörg Hanisch: tilfen@gmail.com" & vbCrLf & "Version 1.16 , 18-01-2012", vbInformation, "About the author"
If mnuchkFloat.Checked Then FloatWindow frmMain.hwnd, True
End Sub

Private Sub mnuchkFloat_Click()
 mnuchkFloat.Checked = IIf(mnuchkFloat.Checked, False, True)
 FloatWindow frmMain.hwnd, mnuchkFloat.Checked
End Sub

Private Sub mnuDonate_Click()
FloatWindow Me.hwnd, False
 frmDonate.Show
End Sub

Private Sub mnuExit_Click()

Dim result
 
 cmddisconnect_Click
 
If ScopeConnectedAtStartup Then

 On Error Resume Next   'Do not EXIT Without Warning!!
 If scope.Connected Then
  FloatWindow frmMain.hwnd, False
    result = MsgBox("ACP will be disconnected from the scope" & vbCrLf & vbCrLf & "Click YES if this causes no problems." & vbCrLf & _
    "Click NO to prepare ACP that it is safe to disconnect", vbYesNo + vbQuestion, "Warning about disconnecting..")
    If result = 7 Then Exit Sub
    'FloatWindow frmMain.hwnd, True
 End If
 
End If

    On Error Resume Next
    If scope.Connected Then scope.Connected = False
    Set scope = Nothing
    Unload Me
 
End Sub

Private Sub mnuExpTS_Click()
 Call SaveTheSkyHorizon
End Sub

Private Sub mnuGetTS_Click()
 Call GetTheSkyHorizon
End Sub


Private Sub mnuHelp_Click()
Dim fs As FileSystemObject
Set fs = New FileSystemObject

Dim HFile As String

    HFile = App.Path & "\AHE.chm"
     If Not fs.FileExists(HFile) Then
     FloatWindow frmMain.hwnd, False
        MsgBox "Help is not available" & vbCrLf & "Please check your files", vbInformation, "Helpfile not found"
        Exit Sub
    Else
        Call HtmlHelp(0, HFile, HH_DISPLAY_TOPIC, ByVal 0&)
    End If
    If mnuchkFloat.Checked Then FloatWindow frmMain.hwnd, True
 Set fs = Nothing
End Sub



Private Sub picHorizon_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

    'Show Azimuth and Elevation
    lblHoriz.Caption = (Round((x / 2), 0) * 2) & "°"
    lblElev.Caption = Format(90 - Y, "0.0") & "°"
    
    picHorizon_MouseMove Button, Shift, x, Y
        
   
End Sub

Private Sub picHorizon_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    
   If x >= 0 And x < 360 Then            'Only draw inside the Box
    If Button = 1 Or Button = 2 Then
        lblHoriz.Caption = Round((x / 2), 0) * 2 & "°"
        lblElev.Caption = Format(90 - Y, "0.0") & "°"
      End If
      
    If Button = 1 Then                'Only proceed if left Mouse-Button is clicked

          
        If Round((x / 2), 0) = 180 Then     'Special treatment (360° = 0°..)
            GridHoriz.TextMatrix(1, 1) = Format(90 - Y, "0.0")
        Else
            GridHoriz.TextMatrix(Round((x / 2), 0) + 1, 1) = Format(90 - Y, "0.0")
        End If
        
        If Round((x / 2), 0) - 11 < 1 Then
            GridHoriz.TopRow = 1
        Else
        GridHoriz.TopRow = Round((x / 2), 0) - 11
        End If
        
        
       
       DrawHorizon
    
    End If
    
  End If

End Sub

Public Sub DrawHorizon()

'If isPaused Then Exit Sub
    
    picHorizon.Refresh
   
   For x = 1 To 180
   On Error Resume Next
        picHorizon.PSet (GridHoriz.TextMatrix(x, 0), 90 - GridHoriz.TextMatrix(x, 1)), QBColor(0)
        'First, draw a blue line from 0 to 90, then replace it yellow
        
        picHorizon.DrawWidth = 3    'needed to cover gaps of 1 degree
        picHorizon.Line (GridHoriz.TextMatrix(x, 0), 90)-(GridHoriz.TextMatrix(x, 0), 0), &H8000000D
        picHorizon.Line (GridHoriz.TextMatrix(x, 0), 90)-(GridHoriz.TextMatrix(x, 0), 90 - GridHoriz.TextMatrix(x, 1)), &H80FFFF   'QBColor(14)
        
        picHorizon.DrawWidth = 2    'reset for Pset only
    Next x
    
CheckHorizon (True)
 
End Sub

'================================================================
'=======Subs for allowing Text-Enty into Flexgrid ==============
'================================================================
Private Sub GridHoriz_Click()
  Call SizeText
  txtInput.Visible = True
  txtInput.SetFocus
  txtInput.Text = GridHoriz.Text
End Sub

Private Sub GridHoriz_GotFocus()
  GridHoriz_RowColChange
End Sub

Private Sub GridHoriz_LeaveCell()

If GridHoriz.Col = 0 Then Exit Sub
  If Not txtInput.Text = "" Then
    GridHoriz.Text = Format(txtInput.Text, "0.0")
    lblHoriz.Caption = (GridHoriz.Row - 1) * 2 & " °"
    lblElev.Caption = Format(txtInput.Text, "0.0") & " °"
  End If
  
End Sub

Private Sub GridHoriz_RowColChange()

  Static OldRow%, OldCol%, Change As Boolean
   If GridHoriz.Col = 0 Then Exit Sub
   

    If Change Then Exit Sub
        Change = True
         
    With GridHoriz
      If .Col <> OldCol Or .Row <> OldRow Then
        OldRow = .Row
        OldCol = .Col
        
        Call SizeText
        txtInput.Visible = True
        txtInput.SetFocus
        txtInput.SelStart = 0
        txtInput.SelLength = Len(txtInput)
      End If
    End With
  Change = False
  
  If Not txtInput.Text = "" Then
    GridHoriz.Text = Format(txtInput.Text, "0.0")
    lblHoriz.Caption = (GridHoriz.Row - 1) * 2 & " °"
    lblElev.Caption = Format(txtInput.Text, "0.0") & " °"
  End If
  
    
  DrawHorizon
  
End Sub




Private Sub txtInput_KeyDown(KeyCode As Integer, Shift As Integer)
  With GridHoriz
    Select Case KeyCode
    
      Case vbKeyRight
        If .Col + 2 > .Cols And .Row + 1 < .Rows Then
          .Col = 1
          .Row = .Row + 1
        ElseIf .Col + 1 < .Cols And .Row < .Rows Then
         .Col = .Col + 1
        End If
      
      Case vbKeyUp
        If .Row - 1 > 0 Then .Row = .Row - 1
      
      Case vbKeyDown, vbKeyReturn
        If .Row + 1 < .Rows Then .Row = .Row + 1
      
      Case vbKeyLeft
        If .Col - 1 = 0 And .Row - 1 <> 0 Then
          .Col = .Cols - 1
          .Row = .Row - 1
        ElseIf .Col - 1 <> 0 Then
          .Col = .Col - 1
        End If
      End Select
    End With
    GridHoriz_RowColChange
End Sub

Private Sub txtInput_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    
    CheckHorizon (True)
  KeyAscii = 0
  End If
  
  If KeyAscii = vbKeyTab Then
    Call txtInput_KeyDown(vbKeyRight, 0)
    KeyAscii = 0
  End If
End Sub

Private Sub txtInput_LostFocus()
  txtInput.Visible = False
  Call GridHoriz_RowColChange

End Sub

Private Sub SizeText()
  With GridHoriz
    txtInput.Text = .Text
    txtInput.FontSize = .Font.Size
    txtInput.Height = .CellHeight
    
    If .CellLeft + .CellWidth > .Width Then
     txtInput.Width = .Width - .CellLeft
    Else
      txtInput.Width = .CellWidth
    End If
    
    txtInput.Left = .CellLeft + .Left
    txtInput.Top = .CellTop + .Top
  End With
End Sub


Private Function CheckHorizon(ByVal singlecheck As Boolean)
Dim val1, val2, val3, val4, val5
Dim counter As Byte
Dim result As Byte
FloatWindow Me.hwnd, False

counter = 0


GridMark.Col = 0
'GridMark.Row = GridHoriz.Row - 1


For x = 1 To 180 '1 To 180

    If x = 1 Then
        val1 = IIf(IsNumeric(frmMain.GridHoriz.TextMatrix(180, 1)), frmMain.GridHoriz.TextMatrix(180, 1), CDbl("0.0"))
    Else
        val1 = IIf(IsNumeric(frmMain.GridHoriz.TextMatrix(x - 1, 1)), frmMain.GridHoriz.TextMatrix(x - 1, 1), CDbl("0.0"))
    End If

        val2 = IIf(IsNumeric(frmMain.GridHoriz.TextMatrix(x, 1)), frmMain.GridHoriz.TextMatrix(x, 1), CDbl("0.0"))

    If x = 180 Then
        val3 = IIf(IsNumeric(frmMain.GridHoriz.TextMatrix(1, 1)), frmMain.GridHoriz.TextMatrix(1, 1), CDbl("0.0"))
    Else
        val3 = IIf(IsNumeric(frmMain.GridHoriz.TextMatrix(x + 1, 1)), frmMain.GridHoriz.TextMatrix(x + 1, 1), CDbl("0.0"))
    End If

        val4 = Abs(val2 - val1)
  
        val5 = Abs(val3 - val2)
  
        GridMark.Row = x - 1
  
    If val4 > 5 Or val5 > 5 Then
        GridMark.CellBackColor = &HC0&
        counter = counter + 1
    Else
        GridMark.CellBackColor = &H8000&
    End If
  
  
Next x

If counter > 0 Then
 
 If singlecheck = False Then
    result = MsgBox("There are two or more horizon points that are over 5 degrees from" & vbCrLf & _
    "their neighbors." & vbCrLf & vbCrLf & "Please click " & Chr(34) & "Cancel" & Chr(34) & " to correct the marked Values in the Table!" & vbCrLf & _
    "Or click " & Chr(34) & "Retry" & Chr(34) & " to save the Horizon and correct it later.", vbRetryCancel + vbExclamation, "Horizon Error...")
        
    CheckHorizon = IIf(result = 2, False, True)
    Else
    CheckHorizon = False
 
 End If

Else
 CheckHorizon = True

End If

If mnuchkFloat.Checked Then FloatWindow Me.hwnd, True
End Function
'================================================================
'=======End of "Text-Entry"-Subs =================================
'================================================================

