VERSION 5.00
Begin VB.Form Login 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Wacko`s PSC Login"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3855
   ClipControls    =   0   'False
   Icon            =   "Login.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   3855
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton grpContestChoice 
      Caption         =   "Power Launcher (tm) "
      Height          =   255
      Index           =   2
      Left            =   960
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   3000
      Width           =   1935
   End
   Begin VB.OptionButton grpContestChoice 
      Caption         =   "World Wide Web Help Wizard(tm)"
      Height          =   255
      Index           =   1
      Left            =   960
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2640
      Width           =   2775
   End
   Begin VB.OptionButton grpContestChoice 
      Caption         =   "Help Maker Plus (tm) "
      Height          =   255
      Index           =   0
      Left            =   960
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2280
      Value           =   -1  'True
      Width           =   1935
   End
   Begin VB.CheckBox chkFeedbackOnComments 
      Caption         =   "email Me when comments are posted"
      Height          =   375
      Left            =   960
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   4320
      Value           =   1  'Checked
      Width           =   2775
   End
   Begin VB.CheckBox chkFreeExposure 
      Caption         =   "I want Free Exposure"
      Height          =   375
      Left            =   960
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3360
      Value           =   1  'Checked
      Width           =   2775
   End
   Begin VB.TextBox txtWWWSite 
      Height          =   285
      Left            =   960
      TabIndex        =   7
      Top             =   3840
      Width           =   2775
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Top             =   1440
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox txtEmailAddress 
      Height          =   285
      Left            =   960
      TabIndex        =   4
      Top             =   1920
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.ComboBox cboWorld 
      Height          =   315
      ItemData        =   "Login.frx":0CCA
      Left            =   960
      List            =   "Login.frx":0CCC
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   0
      Width           =   2775
   End
   Begin VB.CommandButton Exit 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   720
      TabIndex        =   10
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Continue 
      Caption         =   "OK"
      Height          =   375
      Left            =   2040
      TabIndex        =   9
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox txtUser 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   480
      Width           =   2775
   End
   Begin VB.TextBox txtPass 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   960
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   960
      Width           =   2775
   End
   Begin VB.CheckBox New 
      Alignment       =   1  'Right Justify
      Caption         =   "New"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "E-Mail"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   1920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Real Name"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   1440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Web URL"
      Height          =   255
      Left            =   0
      TabIndex        =   16
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "PassWord"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "UserName"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   480
      Width           =   855
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim valContestChoice As String
Dim valFreeExposure As String
Dim ValFeedbackOnComments As String

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Function Win32Keyword(ByVal URL As String) As Long
    weburl = ShellExecute(0&, vbNullString, URL, vbNullString, vbNullString, vbNormalFocus)
End Function

Private Sub chkFeedbackOnComments_Click()
Select Case chkFeedbackOnComments
    Case 0
        ValFeedbackOnComments = ""
    Case 1
        ValFeedbackOnComments = "ON"
End Select
End Sub

Private Function SpaceExchange(Incoming As String) As String
Dim I As Integer
Dim Buffer As String
    For I = 1 To Len(Incoming)
        If Mid(Incoming, I, 1) = " " Then
            Buffer = Mid(Incoming, I + 1, Len(Incoming) - I + 1)
            SpaceExchange = SpaceExchange & Mid(Incoming, 1, I - 1) & "+"
        End If
    Next I
SpaceExchange = SpaceExchange & Buffer
End Function

Private Sub chkFreeExposure_Click()
Select Case chkFreeExposure.Value
    Case 0
        valFreeExposure = ""
    Case 1
        valFreeExposure = "ON"
End Select
End Sub

Private Sub Continue_Click()
If cboWorld.ListIndex = 0 Then MsgBox "Please select a World": Exit Sub
If txtUser = "" Then MsgBox "Please Enter Your Username": Exit Sub
If txtPass = "" Then MsgBox "Please Enter Your Username": Exit Sub
If Me.New.Value = 0 Then
'frmBrowser.brwWebBrowser.Navigate "http://www.planetsourcecode.com/vb/authors/existing_author_login.asp?" & "txtUserId=" & txtUser & "&txtPassword=" & txtPass & "&lngWId=" & cboWorld.ListIndex
Win32Keyword ("http://www.planetsourcecode.com/vb/authors/existing_author_login.asp?" & "txtUserId=" & txtUser & "&txtPassword=" & txtPass & "&lngWId=" & cboWorld.ListIndex)
ElseIf Me.New.Value = 1 Then
'frmBrowser.brwWebBrowser.Navigate "http://www.planetsourcecode.com/vb/authors/new_author_login.asp?" & "blnNewAuthor=" & "&txtUserId=" & txtUser & "&txtPassword=" & txtPass & "&txtEmailAddress=" & txtEmailAddress & "&txtName=" & txtName & "&grpContestChoice=" & valContestChoice & "&txtWWWSite=" & txtWWWSite & "&chkFreeExposure=" & valFreeExposure & "&chkFeedbackOnComments=" & ValFeedbackOnComments & "lngWId=" & cboWorld.ListIndex
Win32Keyword ("http://www.planetsourcecode.com/vb/authors/new_author_login.asp?" & "blnNewAuthor=" & "&txtUserId=" & txtUser & "&txtPassword=" & txtPass & "&txtEmailAddress=" & txtEmailAddress & "&txtName=" & txtName & "&grpContestChoice=" & valContestChoice & "&txtWWWSite=" & txtWWWSite & "&chkFreeExposure=" & valFreeExposure & "&chkFeedbackOnComments=" & ValFeedbackOnComments & "lngWId=" & cboWorld.ListIndex)
End If
'frmBrowser.Visible = True
'Me.Visible = False
'Unload Me
End Sub

Private Sub Exit_Click()
End
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Continue_Click
End Sub

Private Sub Form_Load()
cboWorld.AddItem "Choose Your World", 0
cboWorld.AddItem "Visual Basic World", 1
cboWorld.AddItem "Java / Javascript World", 2
cboWorld.AddItem "C / C++ World", 3
cboWorld.AddItem "ASP / VbScript World", 4
cboWorld.ListIndex = 0
valContestChoice = "Help+Maker+Plus"
valFreeExposure = "ON"
ValFeedbackOnComments = "ON"
'frmBrowser.brwWebBrowser.Navigate "http://www.planetsourcecode.com/vb/default.asp?lngWId=1"
End Sub


Private Sub Form_Resize()
If Me.New.Value = 0 Then
    Me.Exit.Top = 1440
    Me.Continue.Top = 1440
    Me.Height = 2370
    Label4.Visible = False
    Label5.Visible = False
    txtName.Visible = False
    txtEmailAddress.Visible = False
    txtName.TabStop = False
    txtEmailAddress.TabStop = False
    grpContestChoice(0).TabStop = False
    txtWWWSite.TabStop = False
    chkFeedbackOnComments.TabStop = False
    chkFreeExposure.TabStop = False
ElseIf Me.New.Value = 1 Then
    txtName.TabStop = True
    txtEmailAddress.TabStop = True
    grpContestChoice(0).TabStop = True
    txtWWWSite.TabStop = True
    chkFeedbackOnComments.TabStop = True
    chkFreeExposure.TabStop = True
    Me.Exit.Top = 4920
    Me.Continue.Top = 4920
    Me.Height = 5895
    Label4.Visible = True
    Label5.Visible = True
    txtName.Visible = True
    txtEmailAddress.Visible = True
End If
End Sub

Private Sub grpContestChoice_Click(Index As Integer)
Select Case Index
    Case 0
        valContestChoice = "Help+Maker+Plus"
    Case 1
        valContestChoice = "WWW+Help+Wizard"
    Case 2
        valContestChoice = "Power+Launcher"
End Select
End Sub

Private Sub New_Click()
'new_author_login.asp?lngWId=1
Form_Resize
End Sub
'new_author_login.asp?lngWId=1
'"hidden" name="blnNewAuthor"
'type="text" size="20" name="txtUserId"
'type="password" size="20" name="txtPassword"
'type="text" size="20" name="txtEmailAddress"
'type="text" size="20" name="txtName"
'type="radio" Name = "grpContestChoice" value="Help Maker Plus" false>Help Maker Plus (tm) <br>
'type="radio" Name = "grpContestChoice" value="WWW Help Wizard" false>World Wide Web Help Wizard(tm) <br>
'type="radio" Name = "grpContestChoice" value="Power Launcher" false>Power Launcher (tm) <br>
'type="text" size="20" name="txtWWWSite"
'type="checkbox" name="chkFreeExposure" value="ON"
'type="checkbox" Name = "chkFeedbackOnComments" value="ON"

