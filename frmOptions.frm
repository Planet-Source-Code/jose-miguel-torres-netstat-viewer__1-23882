VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Options"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5220
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   5220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdKO 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   555
      Left            =   1500
      TabIndex        =   5
      Top             =   1530
      Width           =   1065
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   555
      Left            =   2640
      TabIndex        =   4
      Top             =   1530
      Width           =   1065
   End
   Begin VB.TextBox txtTime 
      Height          =   360
      Left            =   2400
      TabIndex        =   1
      Top             =   870
      Width           =   675
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Seconds"
      Height          =   240
      Left            =   3180
      TabIndex        =   3
      Top             =   930
      Width           =   765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Update each"
      Height          =   240
      Left            =   1230
      TabIndex        =   2
      Top             =   930
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "You can put the timing on netstat statistics to update it the seconds you indicate below..."
      Height          =   585
      Left            =   150
      TabIndex        =   0
      Top             =   90
      Width           =   4965
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdKO_Click()
frmMain.txtTimer = "Disable"
frmMain.Timer1.Enabled = False
Unload Me
End Sub

Private Sub cmdOK_Click()
frmMain.txtTimer = "Update on " & Me.txtTime & " seconds"
frmMain.Timer1.Interval = Val(Me.txtTime) * 1000
frmMain.Timer1.Enabled = True
Unload Me
End Sub

Private Sub txtTime_LostFocus()
If Not IsNumeric(Val(Me.txtTime)) Or Val(Me.txtTime) < 0 Or Val(Me.txtTime) > 60 Then
    MsgBox "Ups!!! The time interval is 0-60 seconds. Try again", vbCritical
    Me.txtTime.SetFocus
    SendKeys "{Home}+{End}"
End If
End Sub
