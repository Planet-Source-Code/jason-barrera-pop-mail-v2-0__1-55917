Attribute VB_Name = "POPMAIL"
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
Public Type Msgs
 cnt As Integer
 pos As Integer
 html() As Boolean
 boundary() As String
 Headers() As String
 mTo() As String
 From() As String
 Subject() As String
 Date() As String
 body() As String
 BodyHTML() As String
 BodyTXT() As String
 Email() As String
 hasAtt() As Boolean
 attFileName() As String
 attFile() As String
 attList() As String
End Type
Public e_msgs As Msgs
Public Function emailParse(ByVal Message As Variant, ByVal Headers As Variant)
''Do Defaults''
e_msgs.cnt = UBound(Message)
ReDim e_msgs.Email(LBound(Message) To UBound(Message))
ReDim e_msgs.Headers(LBound(Message) To UBound(Message))
ReDim e_msgs.body(LBound(Message) To UBound(Message))
ReDim e_msgs.boundary(LBound(Message) To UBound(Message))
ReDim e_msgs.html(LBound(Message) To UBound(Message))
ReDim e_msgs.mTo(LBound(Message) To UBound(Message))
ReDim e_msgs.From(LBound(Message) To UBound(Message))
ReDim e_msgs.Subject(LBound(Message) To UBound(Message))
ReDim e_msgs.Date(LBound(Message) To UBound(Message))
ReDim e_msgs.attFile(LBound(Message) To UBound(Message))
ReDim e_msgs.attFileName(LBound(Message) To UBound(Message))
ReDim e_msgs.hasAtt(LBound(Message) To UBound(Message))
ReDim e_msgs.BodyHTML(LBound(Message) To UBound(Message))
ReDim e_msgs.BodyTXT(LBound(Message) To UBound(Message))
ReDim e_msgs.attList(LBound(Message) To UBound(Message))

'' Dim some Specific Vars
Dim boundary As String
Dim NoBoundary As Boolean
Dim i As Integer
Dim posA As Long
Dim bPos As Long
Dim bPos2 As Long

'' Execute
For i = LBound(Message) To UBound(Message)

  '' parse headers for to, from, subject etc..
 Dim sp() As String
 sp = Split(Headers(i), vbCrLf)
 Dim s As Variant
 
 For Each s In sp
  If Left(s, 4) = "To: " Then
   If Len(e_msgs.mTo(i)) = 0 Then
   e_msgs.mTo(i) = Mid(s, 5)
   End If
  End If
  If Left(s, 6) = "From: " Then
   If Len(e_msgs.From(i)) = 0 Then
   e_msgs.From(i) = Mid(s, 7)
   End If
  End If
  If Left(s, 9) = "Subject: " Then
   If Len(e_msgs.Subject(i)) = 0 Then
   e_msgs.Subject(i) = Mid(s, 10)
   End If
  End If
  If Left(s, 6) = "Date: " Then
   If Len(e_msgs.Date(i)) = 0 Then
   e_msgs.Date(i) = Mid(s, 7)
   End If
  End If
 Next
 
  '' Load headers
  e_msgs.Headers(i) = Headers(i)
  
  '' Load all
  e_msgs.Email(i) = Message(i)
  
  '' Check for valid Format
  posA = InStr(1, Message(i), vbCrLf & vbCrLf)
  If posA > 0 Then
  
'' Get Boundary
   bPos = InStr(1, Headers(i), "boundary=")
   If bPos > 0 Then
    bPos = InStr(bPos, Headers(i), """") + 1
    bPos2 = InStr(bPos, Headers(i), """")
    boundary = Mid(Headers(i), bPos, bPos2 - bPos)
    e_msgs.boundary(i) = boundary
    NoBoundary = False
   Else
    NoBoundary = True
    boundary = "[_NO__BOUNDARY_]"
    e_msgs.boundary(i) = "[_NO__BOUNDARY_]"
   End If
'' End Get Boundary
    
'' Remove Headers from body
   e_msgs.body(i) = Replace(Message(i), Headers(i), "")

'' Do a little fix
  If NoBoundary = True Then
   e_msgs.body(i) = boundary & vbCrLf & _
   "Content-Type: text/plain;" & vbCrLf & vbCrLf & e_msgs.body(i) & vbCrLf & boundary & vbCrLf & vbCrLf
  End If

'' Do some parsing now
Dim NL As Integer
Dim ctype As String
Dim posx() As Long
Dim tmp As Long
Dim bStr As String
Dim aStr As String
Dim Bcnt As Integer
Dim attStr As String
Bcnt = StrCount(e_msgs.body(i), boundary, False)
NL = Len(boundary)
ReDim posx(1 To Bcnt)


'' Get all boundary positions
  tmp = InStr(1, e_msgs.body(i), boundary, vbTextCompare)
  For x = 1 To StrCount(e_msgs.body(i), boundary, False)
   If tmp > 0 Then
     posx(x) = tmp
     tmp = InStr(tmp + NL, e_msgs.body(i), boundary, vbTextCompare)
   End If
  Next x
  
  
'' Now lets get the good stuff
 For x = 1 To UBound(posx)
  bPos = posx(x) + NL
'' Get the content type
  ctype = getCtype(bPos, e_msgs.body(i))
'' Get the actual message
  bStr = getBody(posx(x), e_msgs.body(i), boundary)
  
  Select Case LCase(ctype)
  
    Case "text/plain"
        e_msgs.BodyTXT(i) = bStr
       
    Case "_no_ctype"
    
    Case "message/disposition-notification"
    
    Case "text/html"
        e_msgs.html(i) = True
        e_msgs.BodyHTML(i) = bStr

    Case Else
'' I suppose this is an attachment, so we will check
        If InStr(bPos, e_msgs.body(i), "Content-Transfer-Encoding: base64", vbTextCompare) > 0 Then
         'Well, it is base64 encoded so we can handle it
          If InStr(bPos, e_msgs.body(i), "Content-Disposition: inline;", vbTextCompare) > 0 Or _
           InStr(bPos, e_msgs.body(i), "Content-Disposition: attachment;", vbTextCompare) > 0 Then
            'OK, we do have an attachment, so lets get it
            'Lets get the filename
            bPos = InStr(bPos, e_msgs.body(i), "filename=" & """", vbTextCompare)
            If bPos > 0 Then
             bPos = bPos + 10
             bPos2 = InStr(bPos, e_msgs.body(i), """", vbTextCompare)
             If bPos2 > 0 Then
              e_msgs.attFileName(i) = Mid(e_msgs.body(i), bPos, bPos2 - bPos)
              attStr = attStr & e_msgs.attFileName(i) & ","
              bPos = InStr(bPos2, e_msgs.body(i), vbCrLf & vbCrLf)
              If bPos > 0 Then
               bPos = bPos + 4
               bPos2 = InStr(bPos, e_msgs.body(i), boundary, vbTextCompare)
               If bPos2 > 0 Then
                aStr = Mid(e_msgs.body(i), bPos, bPos2 - bPos)
                bPos = InStrRev(aStr, "=")
                If bPos > 0 Then
                ''''''''''''''''''''''''''''''''''''
                 e_msgs.attFile(i) = Base64Decode(Left(aStr, bPos))
                 SaveMail e_msgs.attFile(i), App.Path & "\SaveMail\" & e_msgs.attFileName(i)
                 e_msgs.hasAtt(i) = True
                ''''''''''''''''''''''''''''''''''''
                End If
               End If
              End If
             End If
            End If
           End If
        End If

        
  End Select
 Next x
 On Error GoTo err:
 If Len(attStr) > 0 Then
 e_msgs.attList(i) = Left(attStr, Len(attStr) - 1)
 End If
  Else
  
   e_msgs.Headers(i) = "INVALID FORMAT"
   e_msgs.body(i) = "INVALID FORMAT"
   e_msgs.Email(i) = Message(i)
  
  End If
Next i
ReDim e_msgs.attFile(0)
Form2.Show
Form1.Caption = "Ready"
Exit Function
err:
MsgBox err.Description
End Function
Public Function SaveMail(ByVal Mail As String, ByVal Path As String)
Dim Handle As Integer
Handle = FreeFile
Open Path For Binary As #Handle
 Put #Handle, , Mail
Close #Handle
End Function
