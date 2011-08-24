VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.UserControl ucCon 
   ClientHeight    =   435
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   435
   ScaleHeight     =   435
   ScaleWidth      =   435
   Begin MSWinsockLib.Winsock wsHTTP 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "ucCon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'    ucCon for Visual Basic 6, a user control designed to make HTTP requests easy
'    Copyright (C) 2010 John Moore & Mike Campbell
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with this program.  If not, see <http://www.gnu.org/licenses/>.

Option Explicit

Private Type tGet
    sURL As String
    sHost As String
    sProxy As String
    sProxyPort As String
    sCookies As String
    sReferrer As String
    bOn As Boolean
End Type
Private Type tPost
    sURL As String
    sHost As String
    sProxy As String
    sProxyPort As String
    sData As String
    sCookies As String
    sReferrer As String
    bOn As Boolean
End Type
Private GetPending As tGet
Private PostPending As tPost

Private bSendGet As Boolean, bSendPost As Boolean

Private sFileName$, sIncome$, iType%, lLength&
Private sLGetURL$, sLGetHost$, sLGetCookies$
Private sAllCookies$

Public Event ConnectionError(lngNumber%, strDescription$)
Public Event ConnectionConnected()
Public Event ConnectionClosed()
Public Event SendingPOST(strData$)
Public Event SendingGET(strData$)
Public Event DataArrival(strData$)
Public Event DataComplete(strData$, strFileName$)
Public Event MovedDataComplete(strData$)
Public Event CookieSet(strName$, strData$)
Public Event StatusUpdate(strPercentage%)

Private Sub AddToLogFile(strData$)
Dim intFF%: intFF = FreeFile
Open (App.Path & "\log_out.txt") For Append As #intFF
    Print #intFF, strData
Close #intFF
End Sub

Private Sub ClearCon()
    Call wsHTTP_Close
End Sub

Private Sub wsHTTP_Close()
    Dim sOut$: sOut = sIncome
    wsHTTP.Close
    sIncome = ""
    iType = 0
    RaiseEvent ConnectionClosed
    If sOut <> "" Then
        Dim sDecoded$
        sDecoded = DecodeChunked(sOut)
        If sDecoded <> "" Then
            RaiseEvent DataComplete(sDecoded, sFileName)
        Else
            RaiseEvent DataComplete(sOut, sFileName)
        End If
    End If
End Sub
Private Sub wsHTTP_Connect()
    RaiseEvent ConnectionConnected
    If GetPending.bOn = True Then
        SendGet GetPending.sURL, GetPending.sHost, GetPending.sCookies, GetPending.sReferrer, GetPending.sProxy, GetPending.sProxyPort
    ElseIf PostPending.bOn = True Then
        SendPost PostPending.sURL, PostPending.sHost, PostPending.sCookies, PostPending.sData, PostPending.sReferrer, PostPending.sProxy, PostPending.sProxyPort
    End If
End Sub

Private Sub wsHTTP_DataArrival(ByVal bytesTotal As Long)
    Dim sData$, sEvent$, sLocation$, sHost$, sTemp$
    Dim setcookies As String
    wsHTTP.GetData sData, vbString
    sIncome = sIncome & sData
    If sIncome = "" Then Exit Sub

    If InStr(1, sIncome, (vbCrLf & vbCrLf)) Then
        Dim sCookieCheck$
        sCookieCheck = sIncome
        CheckCookies sCookieCheck

        If LCase(Left(sIncome, 7)) = "http/1." Then
            sEvent = Mid(sIncome, 1, InStr(1, sIncome, vbNewLine) - 1)
            sEvent = Mid(sEvent, InStr(1, sEvent, " ") + 1)
            sEvent = Mid(sEvent, InStr(1, sEvent, " ") + 1)
            Select Case LCase(sEvent)
            
                Case "continue":
                    sIncome = ""
                    Exit Sub
                    
                Case "object moved", "found", "moved permanently", "moved temporarily":
                    sLocation = Mid(sIncome, InStr(1, LCase(sIncome), "location:") + 10)
                    sLocation = Mid(sLocation, 1, InStr(1, sLocation, vbNewLine) - 1)
                    If InStr(1, sLocation, "://") Then
                        sLocation = Mid(sLocation, InStr(1, sLocation, "://") + 3)
                        sHost = Mid(sLocation, 1)
                        sHost = Mid(sHost, 1, InStr(1, sHost, "/") - 1)
                        sLocation = Mid(sLocation, InStr(1, sLocation, "/"))
                    Else
                        sHost = wsHTTP.RemoteHost
                    End If
                    sIncome = ""
                    wsHTTP.Close
                    SendGet sLocation, sHost, sAllCookies$, GetLastURL
                    Exit Sub
                    
                Case "ok":
                    If InStr(1, LCase(sIncome), "content-length:") Then
                        iType = 1
                        sTemp = Mid(sIncome, InStr(1, LCase(sIncome), "content-length: ") + 16)
                        sTemp = Mid(sTemp, 1, InStr(1, sTemp, vbNewLine) - 1)
                        lLength = Val(sTemp)
                    ElseIf InStr(1, LCase(sIncome), "transfer-encoding: chunked") Then
                        iType = 2
                    Else
                        iType = 3
                    End If
                    sIncome = Mid(sIncome, InStr(1, sIncome, (vbCrLf & vbCrLf)) + 4)
                 '   If InStr(1, LCase(sIncome), "set-cookie:") Then
                 '       a1 = Split(sIncome, "Set-Cookie:")
                 '       a2 = Split(a1(1), vbNewLine)
                 '       setcookies = a2(0)
                 '   End If
            End Select
        End If
    End If
    
    If iType = 3 Then
        If InStr(1, sIncome, "<!-- onRequestEnd -->") Then
            ClearCon
        End If
    ElseIf iType = 1 Then
        If Len(sIncome) >= lLength Then
            ClearCon
        Else
            RaiseEvent StatusUpdate(((Len(sIncome) * 100) / lLength))
        End If
    End If
End Sub

Private Sub wsHTTP_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    RaiseEvent ConnectionError(Number, Description)
End Sub

Public Function SendGet(sURL$, sHost$, sCookies$, Optional sReferrer$, Optional sProxy$, Optional sProxyPort$)
On Error Resume Next

    sAllCookies = sCookies
    Dim sOut$
    If wsHTTP.State = sckConnected Then
        sFileName = Right(sURL, Len(sURL) - 1)
        Do Until InStr(1, sFileName, "/") = False
            DoEvents
            If InStr(1, sFileName, "/") = False Then Exit Do
            sFileName = Mid(sFileName, InStr(1, sFileName, "/") + 1)
        Loop
        sLGetURL = sURL
        sLGetHost = sHost
        sLGetCookies = sCookies
        If sProxy <> "" Then sURL = "http://" & sHost & sURL
        sOut = "GET " & sURL & " HTTP/1.1" & vbCrLf
        sOut = sOut & "Accept: image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, image/png, application/x-shockwave-flash, */*" & vbCrLf
        sOut = sOut & "Accept-Language: en-us" & vbCrLf
        If sCookies <> "" Then
            sOut = sOut & "Cookie: " & sCookies & vbCrLf
        End If
        If sReferrer <> "" Then
            sOut = sOut & "Referer: " & sReferrer & vbCrLf
        ElseIf sLGetURL <> "" Then
            sOut = sOut & "Referer: http://" & sLGetHost & sLGetURL & vbCrLf
        End If
        sOut = sOut & "User-Agent: Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1; SV1)" & vbCrLf
        sOut = sOut & "Connection: close" & vbCrLf
        sOut = sOut & "Cache-Control: no-cache" & vbCrLf
        sOut = sOut & "Host: " & sHost & vbCrLf
        sOut = sOut & vbCrLf
        GetPending.sURL = ""
        GetPending.sHost = ""
        GetPending.sProxy = ""
        GetPending.sProxyPort = ""
        GetPending.sCookies = ""
        GetPending.sReferrer = ""
        GetPending.bOn = False
        RaiseEvent SendingGET(sOut)
        sIncome = ""
        wsHTTP.SendData sOut
    Else
        GetPending.sURL = sURL
        GetPending.sHost = sHost
        GetPending.sProxy = sProxy
        GetPending.sProxyPort = sProxyPort
        GetPending.sCookies = sCookies
        GetPending.sReferrer = sReferrer
        GetPending.bOn = True
        wsHTTP.Close
        If sProxy <> "" Then
            wsHTTP.Connect sProxy, sProxyPort
        Else
            wsHTTP.Connect sHost, 80
        End If
    End If
End Function

Public Function SendPost(sURL$, sHost$, sCookies$, sData$, Optional sReferrer$, Optional sProxy$, Optional sProxyPort$)
On Error Resume Next

    sAllCookies = sCookies
    Dim sOut$
    If wsHTTP.State = sckConnected Then
        sFileName = Right(sURL, Len(sURL) - 1)
        Do Until InStr(1, sFileName, "/") = False
            DoEvents
            If InStr(1, sFileName, "/") = False Then Exit Do
            sFileName = Mid(sFileName, InStr(1, sFileName, "/") + 1)
        Loop
        sLGetURL = sURL
        sLGetHost = sHost
        sLGetCookies = sCookies
        If sProxy <> "" Then sURL = "http://" & sHost & sURL
        sOut = "POST " & sURL & " HTTP/1.1" & vbCrLf
        sOut = sOut & "Accept: image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, application/x-shockwave-flash, */*" & vbCrLf
        sOut = sOut & "Accept-Language: en-us" & vbCrLf
        sOut = sOut & "Content-Type: application/x-www-form-urlencoded" & vbCrLf
        sOut = sOut & "User-Agent: Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1; SV1)" & vbCrLf
        sOut = sOut & "Host: " & sHost & vbCrLf
        sOut = sOut & "Content-Length: " & Len(sData) & vbCrLf
        sOut = sOut & "Connection: close" & vbCrLf
        sOut = sOut & "Cache-Control: no-cache" & vbCrLf
        'If sReferrer <> "" Then
        '    sOut = sOut & "Referer: " & sReferrer & vbCrLf
        'ElseIf sLGetURL <> "" Then
        '    sOut = sOut & "Referer: http://" & sLGetHost & sLGetURL & vbCrLf
        'End If
        If sCookies <> "" Then
            sOut = sOut & "Cookie: " & sCookies & vbCrLf
        End If
        
        sOut = sOut & vbCrLf
        sOut = sOut & sData
    
        PostPending.sURL = ""
        PostPending.sHost = ""
        PostPending.sProxy = ""
        PostPending.sProxyPort = ""
        PostPending.sCookies = ""
        PostPending.sData = ""
        PostPending.sReferrer = ""
        PostPending.bOn = False
        RaiseEvent SendingPOST(sOut)
        sIncome = ""
        'MsgBox sOut
        wsHTTP.SendData sOut
    Else
        PostPending.sURL = sURL
        PostPending.sHost = sHost
        PostPending.sProxy = sProxy
        PostPending.sProxyPort = sProxyPort
        PostPending.sCookies = sCookies
        PostPending.sReferrer = sReferrer
        PostPending.sData = sData
        PostPending.bOn = True
        wsHTTP.Close
        If sProxy <> "" Then
            wsHTTP.Connect sProxy, sProxyPort
        Else
            wsHTTP.Connect sHost, 80
        End If
    End If
End Function

Public Sub CloseConnection()
    sIncome = ""
    iType = 0
    wsHTTP.Close
    RaiseEvent ConnectionClosed
End Sub

Public Function GetIPAddress() As String
    GetIPAddress = wsHTTP.RemoteHostIP
End Function

Public Function GetHost() As String
    GetHost = wsHTTP.RemoteHost
End Function
Public Function GetState() As String
    GetState = wsHTTP.State
End Function

Public Function GetLastURL() As String
    GetLastURL = "http://" & sLGetHost & sLGetURL
End Function

Private Function DecodeChunked(strSource) As String
    On Error GoTo Exitter
    Dim intA&, intB&, strT$, strD$, lnlL&
    strD = strSource
    intA = InStr(1, strD, vbCrLf)
    lnlL = Val("&H" & Mid(strD, 1, InStr(1, strD, vbCrLf) - 1))
    Do Until lnlL = 0
        strT = strT & Mid(strD, intA + 2, lnlL)
        intB = lnlL + intA + 4
        intA = InStr(intB, strD, vbCrLf)
        lnlL = Val("&H" & Mid(strD, intB, intA - intB))
    Loop
    DecodeChunked = strT
    Exit Function
Exitter:
    DecodeChunked = ""
End Function

Private Sub CheckCookies(strData$)
    Dim strCookN$, strCookD$, sArray() As String, iCount%, sTemp() As String, bFound As Boolean
    Do While InStr(1, LCase(strData), "set-cookie: ")
        strData = Mid(strData, InStr(1, LCase(strData), "set-cookie: ") + 12)
        strCookN = Mid(strData, 1, InStr(1, strData, "=") - 1)
        strData = Mid(strData, InStr(1, strData, "=") + 1)
        If InStr(1, strData, ";") Then
            strCookD = Mid(strData, 1, InStr(1, strData, ";") - 1)
            strData = Mid(strData, InStr(1, strData, ";") + 1)
        End If
        
        RaiseEvent CookieSet(strCookN, strCookD)
        If sAllCookies <> "" Then
            If InStr(1, sAllCookies, ";") Then
                sArray = Split(sAllCookies, ";")
                For iCount = 0 To UBound(sArray)
                    sTemp = Split(sArray(iCount), "=")
                    If LCase(sTemp(0)) = LCase(strCookN) Then
                        sArray(iCount) = strCookN & "=" & strCookD
                        bFound = True
                    End If
                Next iCount
            Else
                sTemp = Split(sAllCookies, "=")
                If LCase(sTemp(0)) = LCase(strCookN) Then
                    sAllCookies = strCookN & "=" & strCookD
                End If
            End If
            If bFound = False Then
                If sAllCookies <> "" Then sAllCookies = sAllCookies & ";"
                sAllCookies = sAllCookies & strCookN & "=" & strCookD
            Else
                sAllCookies = ""
                For iCount = 0 To UBound(sArray)
                    If sAllCookies <> "" Then sAllCookies = sAllCookies & ";"
                    sAllCookies = sAllCookies & sArray(iCount)
                Next iCount
            End If
        Else
            sAllCookies = strCookN & "=" & strCookD
        End If
    Loop
End Sub

