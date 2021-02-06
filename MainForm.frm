VERSION 5.00
Begin VB.Form MainForm 
   Caption         =   "Groovy TestAnswer0r"
   ClientHeight    =   3000
   ClientLeft      =   225
   ClientTop       =   825
   ClientWidth     =   4755
   Icon            =   "MainForm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3000
   ScaleWidth      =   4755
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkSetting 
      Caption         =   "&Upper Case"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   2340
      TabIndex        =   4
      Top             =   540
      Width           =   1155
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "&Generate"
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
      Index           =   1
      Left            =   3540
      TabIndex        =   6
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "&Randomize"
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
      Index           =   0
      Left            =   2340
      TabIndex        =   5
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtOutput 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   7
      Top             =   840
      Width           =   4755
   End
   Begin VB.TextBox txtInput 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   1380
      TabIndex        =   3
      Text            =   "3"
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox txtInput 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   1380
      TabIndex        =   1
      Text            =   "5"
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Answer count:"
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
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   540
      Width           =   1050
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Question count:"
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
      TabIndex        =   0
      Top             =   180
      Width           =   1155
      WordWrap        =   -1  'True
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuRandomize 
         Caption         =   "&Randomize"
      End
      Begin VB.Menu mnuGenerate 
         Caption         =   "&Generate"
      End
      Begin VB.Menu mnus534 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuUpperCase 
         Caption         =   "&Upper Case"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Help Contents"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuLicence 
         Caption         =   "&Licence agreement"
      End
      Begin VB.Menu mnus2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strOutput As String
Dim errRes As Integer

Private Sub cmdButton_Click(Index As Integer)
On Error GoTo ErrH

Select Case Index
    Case 0
        Randomize
    Case 1
        Dim i As Integer
        Dim j As Integer
        
        strOutput = ""
        
        For i = 1 To Int(Me.txtInput(0).Text)
            strOutput = strOutput & (i) & ". " & Chr(Round(Rnd * Int(Me.txtInput(1).Text - 1)) + IIf(Me.chkSetting(0).Value, 65, 97)) & vbCrLf
        Next i
        
        Me.txtOutput.Text = strOutput
End Select

Exit Sub
ErrH:
errRes = MsgBox("Error #" & Err.Number & ":" & vbCrLf & vbCrLf & Err.Description, vbCritical + vbApplicationModal + vbAbortRetryIgnore, App.FileDescription)
If errRes = vbRetry Then Resume
If errRes = vbIgnore Then Resume Next
End Sub

Private Sub Form_Load()
On Error GoTo ErrH

GetRegPos Me
Me.txtInput(0).Text = GetRegLong(HKEY_CURRENT_USER, "Software\" & App.CompanyName & "\" & App.Title, "QuestionCount", 5)
Me.txtInput(1).Text = GetRegLong(HKEY_CURRENT_USER, "Software\" & App.CompanyName & "\" & App.Title, "AnswerCount", 3)
Me.chkSetting(0).Value = GetRegLong(HKEY_CURRENT_USER, "Software\" & App.CompanyName & "\" & App.Title, "UpperCase", False)

Exit Sub
ErrH:
errRes = MsgBox("Error #" & Err.Number & ":" & vbCrLf & vbCrLf & Err.Description, vbCritical + vbApplicationModal + vbAbortRetryIgnore, App.FileDescription)
If errRes = vbRetry Then Resume
If errRes = vbIgnore Then Resume Next
End Sub

Private Sub Form_Resize()
On Error Resume Next
Me.txtOutput.Width = Me.ScaleWidth
Me.txtOutput.Height = Me.ScaleHeight - Me.txtOutput.Top
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrH

SetRegPos Me
SaveRegLong HKEY_CURRENT_USER, "Software\" & App.CompanyName & "\" & App.Title, "QuestionCount", Me.txtInput(0).Text
SaveRegLong HKEY_CURRENT_USER, "Software\" & App.CompanyName & "\" & App.Title, "AnswerCount", Me.txtInput(1).Text
SaveRegLong HKEY_CURRENT_USER, "Software\" & App.CompanyName & "\" & App.Title, "UpperCase", Me.chkSetting(0).Value

Exit Sub
ErrH:
errRes = MsgBox("Error #" & Err.Number & ":" & vbCrLf & vbCrLf & Err.Description, vbCritical + vbApplicationModal + vbAbortRetryIgnore, App.FileDescription)
If errRes = vbRetry Then Resume
If errRes = vbIgnore Then Resume Next
End Sub

Private Sub chkSetting_Click(Index As Integer)
Me.mnuUpperCase.Checked = Me.chkSetting(0).Value
End Sub

Private Sub mnuAbout_Click()
On Error Resume Next
'Load frmAbout
'frmAbout.Left = Me.Left + ((Me.Width - frmAbout.Width) / 2)
'frmAbout.Top = Me.Top + ((Me.Height - frmAbout.Height) / 2)
AboutForm.Show vbModal, Me
End Sub

Private Sub mnuClose_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub mnuGenerate_Click()
cmdButton_Click 1
End Sub

Private Sub mnuHelpContents_Click()
HelpForm.Show
End Sub

Private Sub mnuLicence_Click()
LicenceForm.Show
End Sub

Private Sub mnuRandomize_Click()
cmdButton_Click 0
End Sub

Private Sub mnuUpperCase_Click()
Me.mnuUpperCase.Checked = Not Me.mnuUpperCase.Checked
If Me.mnuUpperCase.Checked Then
    Me.chkSetting(0).Value = 1
Else
    Me.chkSetting(0).Value = 0
End If
End Sub

Private Sub txtInput_Change(Index As Integer)
On Error GoTo ErrH

If Val(Me.txtInput(Index).Text) < 1 Then
    Me.txtInput(Index).Text = "1"
End If

If Index = 0 Then
    If Val(Me.txtInput(Index).Text) > 5000 Then
        Me.txtInput(Index).Text = "5000"
    End If
ElseIf Index = 1 Then
    If Val(Me.txtInput(Index).Text) > 26 Then
        Me.txtInput(Index).Text = "26"
    End If
End If

Exit Sub
ErrH:
errRes = MsgBox("Error #" & Err.Number & ":" & vbCrLf & vbCrLf & Err.Description, vbCritical + vbApplicationModal + vbAbortRetryIgnore, App.FileDescription)
If errRes = vbRetry Then Resume
If errRes = vbIgnore Then Resume Next
End Sub
