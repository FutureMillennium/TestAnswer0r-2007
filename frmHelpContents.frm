VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form HelpForm 
   Caption         =   "Help Contents"
   ClientHeight    =   5100
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   7290
   Icon            =   "frmHelpContents.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5100
   ScaleWidth      =   7290
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   5
      Top             =   4785
      Width           =   7290
      _ExtentX        =   12859
      _ExtentY        =   556
      Style           =   1
      SimpleText      =   "Ready"
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtSource 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      Text            =   "frmHelpContents.frx":058A
      Top             =   600
      Visible         =   0   'False
      Width           =   5115
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin SHDocVwCtl.WebBrowser theContent 
      Height          =   1335
      Left            =   0
      TabIndex        =   3
      Top             =   600
      Width           =   1995
      ExtentX         =   3519
      ExtentY         =   2355
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   0
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "http://testanswer0r.justgroovy.net/software-help"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Index           =   1
      Left            =   1440
      TabIndex        =   2
      Top             =   210
      Width           =   3975
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Also available at:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   210
      Width           =   1275
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "HelpForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub WriteHelp()
On Error Resume Next
Me.theContent.Document.body.innerHTML = Me.txtSource.Text
End Sub

Private Sub Form_Load()
On Error GoTo ErrH

GetRegPos Me
Me.Caption = App.ProductName & " Help Contents"
Me.theContent.Navigate2 "about:blank"
Me.lblInfo(1).ToolTipText = "Go to " & Me.lblInfo(1).Caption

Exit Sub
ErrH:
errRes = MsgBox("Error #" & Err.Number & ":" & vbCrLf & vbCrLf & Err.Description, vbCritical + vbApplicationModal + vbAbortRetryIgnore, App.FileDescription)
If errRes = vbRetry Then Resume
If errRes = vbIgnore Then Resume Next
End Sub

Private Sub Form_Resize()
On Error Resume Next
Me.theContent.Width = Me.ScaleWidth
Me.theContent.Height = Me.ScaleHeight - Me.theContent.Top - Me.StatusBar1.Height
Me.cmdClose.Left = Me.ScaleWidth - Me.cmdClose.Width - 120
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrH

SetRegPos Me

Exit Sub
ErrH:
errRes = MsgBox("Error #" & Err.Number & ":" & vbCrLf & vbCrLf & Err.Description, vbCritical + vbApplicationModal + vbAbortRetryIgnore, App.FileDescription)
If errRes = vbRetry Then Resume
If errRes = vbIgnore Then Resume Next
End Sub

Private Sub lblInfo_Click(Index As Integer)
If Left$(Me.lblInfo(Index).Caption, 7) = "http://" Then
    ShellExecute Me.hWnd, vbNullString, Me.lblInfo(Index).Caption, vbNullString, "C:\", SW_SHOWNORMAL
End If
End Sub

Private Sub theContent_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
WriteHelp
End Sub

Private Sub theContent_StatusTextChange(ByVal Text As String)
Me.StatusBar1.SimpleText = Text
End Sub
