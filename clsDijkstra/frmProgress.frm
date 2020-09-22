VERSION 5.00
Begin VB.Form frmProgress 
   BorderStyle     =   0  'None
   ClientHeight    =   1080
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2355
   Icon            =   "frmProgress.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1080
   ScaleWidth      =   2355
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraProgress 
      Caption         =   "Progress"
      Height          =   1050
      Left            =   30
      TabIndex        =   0
      Top             =   -15
      Width           =   2280
      Begin VB.CommandButton cmdStop 
         Caption         =   "Stop"
         Height          =   420
         Left            =   480
         TabIndex        =   1
         Top             =   510
         Width           =   1200
      End
      Begin VB.Label lblProgress 
         BackStyle       =   0  'Transparent
         Caption         =   "0 %"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   195
         TabIndex        =   2
         Top             =   225
         Width           =   450
      End
      Begin VB.Shape shpProgress 
         BackColor       =   &H00000000&
         BorderColor     =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   225
         Left            =   135
         Top             =   210
         Width           =   15
      End
      Begin VB.Shape shpBack 
         BackColor       =   &H8000000F&
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H8000000F&
         FillStyle       =   0  'Solid
         Height          =   225
         Left            =   135
         Top             =   210
         Width           =   1995
      End
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Just a simple "progress bar" form, also shows % in taskbar.

Option Explicit

Dim X1 As Single, Y1 As Single

Private Sub Form_Load()
    Icon = frmMDI.Icon
End Sub

Private Sub cmdStop_Click()
    blStopProcess = True
End Sub

Private Sub cmdStop_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If X1 <> x And Y1 <> y Then
        cmdStop.ToolTipText = ShowTip
        X1 = x
        Y1 = y
    End If
End Sub
