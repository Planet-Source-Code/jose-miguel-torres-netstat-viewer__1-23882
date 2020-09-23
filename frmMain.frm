VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NetStatViewer"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9450
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   9450
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   8460
      Top             =   1830
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":08CA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   9150
      Top             =   570
   End
   Begin VB.TextBox txtNS 
      BackColor       =   &H80000001&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   2505
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   810
      Width           =   9105
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   3045
      Left            =   60
      TabIndex        =   4
      Top             =   390
      Width           =   9225
      _ExtentX        =   16272
      _ExtentY        =   5371
      Style           =   2
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Active Connections"
            Object.ToolTipText     =   "Netstat -n"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Active Connections and Ports"
            Object.ToolTipText     =   "Netstat -a"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Table of routes"
            Object.ToolTipText     =   "Netstat -r"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "IPConfig"
            Object.ToolTipText     =   "ipconfig"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Ping x?"
            Object.ToolTipText     =   "ping"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8820
      Top             =   3570
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0BE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":17D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1AF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2304
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2756
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2BA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3482
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5C34
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5F4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7C58
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8532
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8E0C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   2
      Top             =   6195
      Width           =   9450
      _ExtentX        =   16669
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "NetStatViewer"
            TextSave        =   "NetStatViewer"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   8819
            MinWidth        =   8819
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "12:24  Josep Miquel"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "08/06/2001"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9450
      _ExtentX        =   16669
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList1"
      DisabledImageList=   "ImageList1"
      HotImageList    =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Config"
            Object.ToolTipText     =   "Configuration"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "NSH"
            Object.ToolTipText     =   "Netstat Help"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "NBS"
            Object.ToolTipText     =   "NetBios Help"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "PH"
            Object.ToolTipText     =   "Ping Help"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Inst"
            Object.ToolTipText     =   "My instruction..."
            ImageIndex      =   11
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Contact"
            Object.ToolTipText     =   "Contact Me"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Web"
            Object.ToolTipText     =   "Link NetViewer Internet Site"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "About"
            Object.ToolTipText     =   "About"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Exit"
            ImageIndex      =   6
         EndProperty
      EndProperty
      Begin VB.TextBox txtTimer 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "Disable"
         Top             =   30
         Width           =   1665
      End
      Begin MSComctlLib.ImageCombo ImageCombo1 
         Height          =   345
         Left            =   6870
         TabIndex        =   6
         Top             =   -30
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   609
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
         ImageList       =   "ImageList2"
      End
   End
   Begin VB.TextBox txtNSE 
      BackColor       =   &H80000001&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   2505
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   3570
      Width           =   7995
   End
   Begin MSComctlLib.TabStrip TabStrip2 
      Height          =   2565
      Left            =   60
      TabIndex        =   5
      Top             =   3540
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   4524
      MultiRow        =   -1  'True
      Style           =   2
      Placement       =   3
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "NetStat Statistics"
            Object.ToolTipText     =   "Netstat -e"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "By Protocol"
            Object.ToolTipText     =   "Netstat -es"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "NetBios State"
            Object.ToolTipText     =   "nbtstat -s"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
Option Explicit

'
' This project shows netstat instrution of commnand and show it in Windows Forms
'
' Declare all API, first

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

Private Declare Function WaitForSingleObject Lib "kernel32" _
    (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" (ByVal _
   hObject As Long) As Long

Const SW_SHOWMINNOACTIVE = 7
Const STARTF_USESHOWWINDOW = &H1
Const INFINITE = -1&
Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const STARTF_USESTDHANDLES = &H100&

' to execute the browser
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOWNORMAL = 1
Private Const SW_SHOWMAXIMIZED = 3
Private Const SW_SHOWDEFAULT = 10
'delay function
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Private Function ExecCmdPipe(ByVal CmdLine As String) As String
    'Executes the command, and when it finish returns value to VB

    Dim proc As PROCESS_INFORMATION, ret As Long, bSuccess As Long
    Dim start As STARTUPINFO
    Dim sa As SECURITY_ATTRIBUTES
    Dim hReadPipe As Long, hWritePipe As Long
    Dim bytesread As Long, mybuff As String
    Dim i As Integer
    
    Dim sReturnStr As String
    
    ' the lenght of the string must be 10 * 1024
    
    mybuff = String(10 * 1024, Chr$(65))
    sa.nLength = Len(sa)
    sa.bInheritHandle = 1&
    sa.lpSecurityDescriptor = 0&
    ret = CreatePipe(hReadPipe, hWritePipe, sa, 0)
    If ret = 0 Then
        '===Error
        ExecCmdPipe = "Error: CreatePipe failed. " & Err.LastDllError
        Exit Function
    End If
    start.cb = Len(start)
    start.hStdOutput = hWritePipe
    start.dwFlags = STARTF_USESTDHANDLES + STARTF_USESHOWWINDOW
    start.wShowWindow = SW_SHOWMINNOACTIVE
    
    ' Start the shelled application:
    ret& = CreateProcessA(0&, CmdLine$, sa, sa, 1&, _
        NORMAL_PRIORITY_CLASS, 0&, 0&, start, proc)
    If ret <> 1 Then
        '===Error
        sReturnStr = "Error: CreateProcess failed. " & Err.LastDllError
    End If
    
    ' Wait for the shelled application to finish:
    ret = WaitForSingleObject(proc.hProcess, INFINITE)
    

    bSuccess = ReadFile(hReadPipe, mybuff, Len(mybuff), bytesread, 0&)
    If bSuccess = 1 Then
        sReturnStr = Left(mybuff, bytesread)
    Else
        '===Error
        sReturnStr = "Error: ReadFile failed. " & Err.LastDllError
    End If
    ret = CloseHandle(proc.hProcess)
    ret = CloseHandle(proc.hThread)
    ret = CloseHandle(hReadPipe)
    ret = CloseHandle(hWritePipe)
    
    'returns to VB
    ExecCmdPipe = sReturnStr
End Function


Private Sub Form_Load()
    Show
    'WELCOME

    Me.txtNS.Text = "Welcome to NetStat Viewer. You could view all netstat, nbtstat, ipconfig and ping instructions into a Windows Form."
    
    
    Me.ImageCombo1.ComboItems.Add 1, "netstat -e", "netstat -e", 1, 1
End Sub



Private Sub Form_Unload(Cancel As Integer)
' exit or not
If MsgBox("Exit Netstat Viewer? ", vbYesNo) = vbNo Then
    Cancel = True
End If
End Sub

Private Sub TabStrip1_Click()
Screen.MousePointer = 11
Me.txtNS = "Please wait..."
Me.txtNS.Refresh
If TabStrip1.SelectedItem.Index = 5 Then
    Me.txtNS = "Waiting IP address..."
    Me.txtNS.Refresh
    Me.txtNS = ExecCmdPipe("ping " & InputBox("Enter IP address:", "Ping", "0.0.0.0"))
Else
    Me.txtNS = ExecCmdPipe(TabStrip1.SelectedItem.ToolTipText)
End If
Screen.MousePointer = 0
End Sub

Private Sub TabStrip2_Click()
Dim strInst As String
Screen.MousePointer = 11
Me.txtNSE = "Please wait..."
Me.txtNSE.Refresh
Me.txtNSE = ExecCmdPipe(TabStrip2.SelectedItem.ToolTipText)
Screen.MousePointer = 0

End Sub

Private Sub Timer1_Timer()
    If Me.ImageCombo1.SelectedItem.Text <> "" Then Me.txtNSE = ExecCmdPipe(Me.ImageCombo1.SelectedItem.Text)
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.Index = 1 Then 'config
    frmOptions.txtTime = Me.Timer1.Interval
    frmOptions.Show 1, Me
ElseIf Button.Index = 3 Then 'help
    Me.txtNS = "Allows you to view the netstat command in a Windows Form"
ElseIf Button.Index = 4 Then 'help
    Me.txtNS = "Allows you to view the nbtstat command in a Windows Form"
ElseIf Button.Index = 5 Then 'help
    Me.txtNS = ExecCmdPipe("ping")
ElseIf Button.Index = 7 Then ' myinstruction
    Me.ImageCombo1.ComboItems.Add Me.ImageCombo1.ComboItems.Count + 1, Me.ImageCombo1.ComboItems.Count + 1 & "I", InputBox("Please, enter your instruction:", "My instruction", "Netstat [terminal] -e ..."), 1, 1
ElseIf Button.Index = 8 Then '
    MsgBox "Its Free button"
ElseIf Button.Index = 10 Then 'contactme
    ShellExecute frmMain.hwnd, "Open", "mailto:jtorres_diaz@teleline.es?subject=About NetViewer", 0&, 0&, SW_SHOWMINNOACTIVE
ElseIf Button.Index = 11 Then 'web site
    ShellExecute frmMain.hwnd, "Open", "http://www.geocities.com/netviewer2002", 0&, 0&, SW_SHOWMINNOACTIVE
ElseIf Button.Index = 13 Then 'about
    frmAbout.Show 1, Me
ElseIf Button.Index = 15 Then 'exit
    Unload Me
End If
End Sub

