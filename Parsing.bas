Attribute VB_Name = "Parsing"
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
Public Function getCtype(ByVal posx As Long, body As String) As String

Dim pos1 As Long
Dim pos2 As Long
Dim line As String

pos1 = InStr(posx, body, "Content-Type: ")
If pos1 Then
pos1 = pos1 + 14
pos2 = InStr(pos1, body, vbCrLf)
  If pos2 Then
   line = Mid(body, pos1, pos2 - pos1)
   Else
   line = "text/plain"
  End If
  Else
  pos2 = InStr(posx, body, vbCrLf & vbCrLf)
  If pos2 Then
   line = "_NO_CTYPE"
   Else
   line = "_NO_CTYPE"
  End If
End If
If line <> "_NO_CTYPE" Then

pos1 = 1
pos2 = InStr(pos1, line, ";")
If pos2 Then
 line = Mid(line, pos1, pos2 - pos1)
End If
line = Replace(line, " ", "")
End If
getCtype = line
End Function
Public Function getBody(posx As Long, body As String, boundary As String) As String
Dim pos1 As Long
Dim pos2 As Long
Dim str As String

pos1 = InStr(posx, body, vbCrLf & vbCrLf, vbTextCompare)
If pos1 Then
 pos1 = pos1 + 2
 pos2 = InStr(pos1, body, boundary, vbTextCompare)
 If pos2 Then
  str = Mid(body, pos1 + 2, (pos2 - pos1) - 4)
 End If
End If
getBody = str
End Function
