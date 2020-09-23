VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   6375
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7725
   LinkTopic       =   "Form2"
   ScaleHeight     =   6375
   ScaleWidth      =   7725
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox cboAtt 
      Height          =   255
      ItemData        =   "Form2.frx":0000
      Left            =   2880
      List            =   "Form2.frx":0002
      TabIndex        =   15
      Top             =   1680
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox txtDATE 
      Height          =   285
      Left            =   600
      TabIndex        =   12
      Top             =   1320
      Width           =   2055
   End
   Begin VB.TextBox txtSUBJECT 
      Height          =   285
      Left            =   3360
      TabIndex        =   10
      Top             =   1320
      Width           =   4215
   End
   Begin VB.TextBox txtFROM 
      Height          =   285
      Left            =   4200
      TabIndex        =   8
      Top             =   960
      Width           =   3375
   End
   Begin VB.TextBox txtTO 
      Height          =   285
      Left            =   480
      TabIndex        =   6
      Top             =   960
      Width           =   3135
   End
   Begin SHDocVwCtl.WebBrowser html 
      Height          =   3975
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   7455
      ExtentX         =   13150
      ExtentY         =   7011
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
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
      Location        =   "http:///"
   End
   Begin VB.TextBox txtHeaders 
      Height          =   855
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   0
      Width           =   7455
   End
   Begin VB.CommandButton Command2 
      Caption         =   ">>"
      Height          =   255
      Left            =   3960
      TabIndex        =   2
      Top             =   6120
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<<"
      Height          =   255
      Left            =   3360
      TabIndex        =   1
      Top             =   6120
      Width           =   495
   End
   Begin VB.TextBox txtMessages 
      Height          =   3975
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   2040
      Width           =   7455
   End
   Begin VB.Label Label5 
      Caption         =   "Attachments:"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label lblAtt 
      Height          =   255
      Left            =   1080
      TabIndex        =   13
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Date:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Subject:"
      Height          =   255
      Left            =   2760
      TabIndex        =   9
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "From:"
      Height          =   255
      Left            =   3720
      TabIndex        =   7
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "To:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   375
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Author: Jason Barrera
'' WWW: http://www.cybercleveland.com
'' CopyRight 2004
'' If you use this code please give credit where
'' credit is due! Thanks..
'''''''''''''''''''''''''
'' I got the POP session idea from vbip.com,
'' which has been mostly all been rewritten
'''''''''''''''''''''''''
'' All Base64 code was done by someone but dont remember
'' who, but I got it from psc.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
        (ByVal hwnd As Long, ByVal lpOperation As String, _
        ByVal lpFile As String, ByVal lpParameters As String, _
        ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Type ControlPositionType
    Left As Single
    Top As Single
    Width As Single
    Height As Single
    FontSize As Single
End Type

Private m_ControlPositions() As ControlPositionType
Private m_FormWid As Single
Private m_FormHgt As Single


Dim WithEvents WB As HTMLDocument
Attribute WB.VB_VarHelpID = -1
Dim Attachments As String
Private Sub ResizeControls()
Dim i As Integer
Dim ctl As Control
Dim x_scale As Single
Dim y_scale As Single

    ' Don't bother if we are minimized.
    If WindowState = vbMinimized Then Exit Sub

    ' Get the form's current scale factors.
    x_scale = ScaleWidth / m_FormWid
    y_scale = ScaleHeight / m_FormHgt

    ' Position the controls.
    i = 1
    For Each ctl In Controls
        With m_ControlPositions(i)
            If TypeOf ctl Is line Then
                ctl.X1 = x_scale * .Left
                ctl.Y1 = y_scale * .Top
                ctl.X2 = ctl.X1 + x_scale * .Width
                ctl.Y2 = ctl.Y1 + y_scale * .Height
            Else
                ctl.Left = x_scale * .Left
                ctl.Top = y_scale * .Top
                ctl.Width = x_scale * .Width
                If Not (TypeOf ctl Is ComboBox) Then
                    ' Cannot change height of ComboBoxes.
                    ctl.Height = y_scale * .Height
                End If
                On Error Resume Next
                ctl.Font.Size = y_scale * .FontSize
                On Error GoTo 0
            End If
        End With
        i = i + 1
    Next ctl
End Sub
Private Sub SaveSizes()
Dim i As Integer
Dim ctl As Control

    ' Save the controls' positions and sizes.
    ReDim m_ControlPositions(1 To Controls.Count)
    i = 1
    For Each ctl In Controls
        With m_ControlPositions(i)
            If TypeOf ctl Is line Then
                .Left = ctl.X1
                .Top = ctl.Y1
                .Width = ctl.X2 - ctl.X1
                .Height = ctl.Y2 - ctl.Y1
            Else
                .Left = ctl.Left
                .Top = ctl.Top
                .Width = ctl.Width
                .Height = ctl.Height
                On Error Resume Next
                .FontSize = ctl.Font.Size
                On Error GoTo 0
            End If
        End With
        i = i + 1
    Next ctl

    ' Save the form's size.
    m_FormWid = ScaleWidth
    m_FormHgt = ScaleHeight
End Sub

Private Sub Command1_Click()
If e_msgs.pos = 1 Then
 Exit Sub
Else
  If e_msgs.hasAtt(e_msgs.pos - 1) = True Then
   cboAtt.Clear
   Attachments = e_msgs.attList(e_msgs.pos - 1)
   lblAtt = (StrCount(Attachments, ",", False) + 1) & " attachments!"
   LoadList cboAtt, Attachments
   Else
   Attachments = ""
   lblAtt = "0 attachments!"
   cboAtt.Visible = False
   cboAtt.Clear
  End If
 txtHeaders = e_msgs.Headers(e_msgs.pos - 1)
 txtMessages = e_msgs.BodyTXT(e_msgs.pos - 1)
 txtTO = e_msgs.mTo(e_msgs.pos - 1)
 txtFROM = e_msgs.From(e_msgs.pos - 1)
 txtSUBJECT = e_msgs.Subject(e_msgs.pos - 1)
 txtDATE = e_msgs.Date(e_msgs.pos - 1)
 e_msgs.pos = e_msgs.pos - 1
 Me.Caption = "Viewing #" & e_msgs.pos & " of " & e_msgs.cnt & " email messages"
If e_msgs.html(e_msgs.pos) = True Then
html.Visible = True
Do Until html.ReadyState = READYSTATE_COMPLETE
DoEvents
Loop
WB.body.innerHTML = e_msgs.BodyHTML(e_msgs.pos)
Else
html.Visible = False
End If
End If
End Sub

Private Sub Command2_Click()
 If e_msgs.pos = e_msgs.cnt Then
  Exit Sub
  Else
  If e_msgs.hasAtt(e_msgs.pos + 1) = True Then
   cboAtt.Clear
   Attachments = e_msgs.attList(e_msgs.pos + 1)
   lblAtt = (StrCount(Attachments, ",", False) + 1) & " attachments!"
   LoadList cboAtt, Attachments
   Else
   Attachments = ""
   lblAtt = "0 attachments!"
   cboAtt.Visible = False
   cboAtt.Clear
  End If
  txtHeaders = e_msgs.Headers(e_msgs.pos + 1)
  txtMessages = e_msgs.BodyTXT(e_msgs.pos + 1)
  txtTO = e_msgs.mTo(e_msgs.pos + 1)
  txtFROM = e_msgs.From(e_msgs.pos + 1)
  txtSUBJECT = e_msgs.Subject(e_msgs.pos + 1)
  txtDATE = e_msgs.Date(e_msgs.pos + 1)
  e_msgs.pos = e_msgs.pos + 1
  Me.Caption = "Viewing #" & e_msgs.pos & " of " & e_msgs.cnt & " email messages"
 If e_msgs.html(e_msgs.pos) = True Then
html.Visible = True
Do Until html.ReadyState = READYSTATE_COMPLETE
DoEvents
Loop
WB.body.innerHTML = e_msgs.BodyHTML(e_msgs.pos)
Else
html.Visible = False
End If
 End If
End Sub



Private Sub Form_Load()
On Error Resume Next
SaveSizes
html.Visible = False
html.Navigate "about:blank"
Set WB = html.Document
e_msgs.pos = 1
  If e_msgs.hasAtt(1) = True Then
   cboAtt.Clear
   Attachments = e_msgs.attList(1)
   lblAtt = (StrCount(Attachments, ",", False) + 1) & " attachments!"
   LoadList cboAtt, Attachments
   Else
   Attachments = ""
   lblAtt = "0 attachments!"
   cboAtt.Visible = False
   cboAtt.Clear
  End If
txtHeaders = e_msgs.Headers(1)
txtMessages = e_msgs.BodyTXT(1)
txtTO = e_msgs.mTo(1)
txtFROM = e_msgs.From(1)
txtSUBJECT = e_msgs.Subject(1)
txtDATE = e_msgs.Date(1)
Me.Caption = "Viewing #1 of " & e_msgs.cnt & " email messages"
If e_msgs.html(1) = True Then
html.Visible = True
On Error Resume Next
Do Until html.ReadyState = READYSTATE_COMPLETE
DoEvents
Loop
WB.body.innerHTML = e_msgs.BodyHTML(1)
Else
html.Visible = False
End If
End Sub

Private Sub Form_Resize()
ResizeControls
If Me.WindowState = 0 Then
'' 255
cboAtt.Height = 255
ElseIf Me.WindowState = 2 Then
cboAtt.Height = lblAtt.Height
End If
End Sub

Private Sub html_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
If InStr(1, URL, "http") > 0 Or InStr(1, URL, "#") > 0 Then
 Cancel = True
End If
End Sub

Private Sub html_DocumentComplete(ByVal pDisp As Object, URL As Variant)
'WB.Body.innerHTML = ""
End Sub
Private Sub WB_onmousedown()
Dim eventObj As IHTMLEventObj
Set eventObj = WB.parentWindow.event
Dim r As Long
If eventObj.button = 1 Then
 If LCase(eventObj.srcElement.tagName) = "a" Then
  If InStr(1, eventObj.srcElement, "http") > 0 Then

   r = ShellExecute(Me.hwnd, "Open", Mid(eventObj.srcElement, InStr(1, eventObj.srcElement, "http")), _
       vbNullString, App.Path, vbNormalFocus)

  End If
 End If
End If
If eventObj.button = 2 Then
 If LCase(eventObj.srcElement.tagName) = "a" Then
  If InStr(1, eventObj.srcElement, "http") > 0 Then

   r = ShellExecute(Me.hwnd, "Open", Mid(eventObj.srcElement, InStr(1, eventObj.srcElement, "http")), _
       vbNullString, App.Path, vbNormalFocus)

  End If
 End If
End If
End Sub
Private Function LoadList(cbo As ListBox, str As String)
Dim s() As String
Dim name As Variant
cbo.Visible = True
If InStr(1, str, ",") > 0 Then
s = Split(str, ",")
i = 1
For Each name In s
cbo.AddItem name
Next
Else
cbo.AddItem str
End If
End Function
