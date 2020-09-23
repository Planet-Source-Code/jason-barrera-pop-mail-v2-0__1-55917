Attribute VB_Name = "Misc"
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
Public Function SockStatus(ByVal sock As Winsock) As String
  Dim strMessage As String
    Select Case sock.State

        Case StateConstants.sckConnected
            strMessage = "Connected to " & sock.RemoteHost
        Case StateConstants.sckClosing
            strMessage = "Closing connection to " & sock.RemoteHost
        Case StateConstants.sckClosed
            strMessage = "Not Connected"
        Case StateConstants.sckError
            strMessage = "Error in Socket"
        Case StateConstants.sckConnected
            strMessage = "Connecting to " & sock.RemoteHost
        Case StateConstants.sckHostResolved
            strMessage = sock.RemoteHost & " Resolved"
        Case StateConstants.sckOpen
            strMessage = "Opened Socket"
        Case StateConstants.sckResolvingHost
            strMessage = "Resolving " & sock.RemoteHost
        Case StateConstants.sckConnectionPending
            strMessage = "Connection is in Pending"
        Case StateConstants.sckListening
            strMessage = "Awaiting connection Request"

    End Select

SockStatus = strMessage
End Function

