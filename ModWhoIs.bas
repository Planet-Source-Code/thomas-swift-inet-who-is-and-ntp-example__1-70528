Attribute VB_Name = "ModWhoIs"
'Took me the better part of 4 hours to debug this little jewl
'But with no more ferther adue, Its the INET WhoIs ! And now INET NTP.
'By:Thomas A. Swift May 12,2008
Option Explicit
Public Function WhoIs(dInet As Inet, IP As String, Optional Host As String = "whois.arin.net", Optional Port As String = "43", Optional Timeout As Long = 60) As String
    dInet.RemoteHost = Host
    dInet.RemotePort = Port
    dInet.RequestTimeout = Timeout
    dInet.AccessType = icDirect
    Call dInet.Execute(vbNullString, IP) 'Submit the IP address
    Do Until dInet.StillExecuting = False
        DoEvents
    Loop
    WhoIs = Replace(dInet.GetChunk(65535, icString), Chr(10), vbCrLf)
    dInet.Cancel 'Same as close of winsock
End Function
'Took me the rest of the night to get this one working.
Public Function GetTime(dInet As Inet, Optional Host As String = "time-a.nist.gov", Optional Port As String = "37", Optional Timeout As Long = 60) As Date
    Dim sTemp As String
    Dim dblNTPTime As Double
    Dim LngTimeFrom1990 As Long
    Dim sngTimeDelay As Single
    dInet.RemoteHost = Host
    dInet.RemotePort = Port
    dInet.RequestTimeout = Timeout
    dInet.AccessType = icDirect
    dInet.Execute 'Get the time
    Do Until dInet.StillExecuting = False
        DoEvents
    Loop
    sTemp = Trim(dInet.GetChunk(1024, icString))
    If Len(sTemp) <> 4 Then
        Call MsgBox("NTP Server returned an invalid response.", _
                vbCritical, "Invalid Response")
        Exit Function
    End If
    dblNTPTime = Asc(Left$(sTemp, 1)) * 256 ^ 3 + Asc(Mid$(sTemp, 2, 1)) * 256 ^ 2 + Asc(Mid$(sTemp, 3, 1)) * 256 ^ 1 + Asc(Right$(sTemp, 1))
    LngTimeFrom1990 = dblNTPTime - 2840140800#
    GetTime = DateAdd("s", CDbl(LngTimeFrom1990 + CLng(sngTimeDelay)), #1/1/1990#)
    dInet.Cancel 'Same as close of winsock
End Function
