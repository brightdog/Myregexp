VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsXMLHTTPGetHtml"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub Sleep _
                Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim strURL As String
Dim strMyPostData As String
Dim strCharSet As String
Dim iRetry As Integer
Private strReferer As String


Dim dicCookies As Scripting.Dictionary

Private dicRequestHeader As Scripting.Dictionary
Private Const iSleepTime = 2

Public Function StartGetHtml() As String
        '<EhHeader>
        On Error GoTo StartGetHtml_Err
        '</EhHeader>

100     If strURL = "" Then

102         StartGetHtml = ""
            Exit Function

        End If
    
104     iRetry = 0

        Dim oReq As New MSXML2.XMLHTTP
    
Retry:

106     If strMyPostData = "" Then

108         oReq.Open "GET", URL, True
            '        oReq.setRequestHeader "Accept", "Text/Html"
            '        oReq.setRequestHeader "Accept-Encoding", "deflate"
            '        oReq.setRequestHeader "Accept-Charset", "GB2312"
            'oReq.Open "Get", "http://" & Url, True
110         Call SetCookies(oReq)
112         oReq.send Null

        Else

114         oReq.Open "POST", URL, True
116         oReq.SetRequestHeader "Content-Length", Len(strMyPostData)
118         oReq.SetRequestHeader "CONTENT-TYPE", "application/x-www-form-urlencoded"
120         oReq.SetRequestHeader "Accept", "Text/Html"
122         oReq.SetRequestHeader "Accept-Encoding", "deflate"
124         Call SetCookies(oReq)
126         oReq.send strMyPostData

        End If

        'oreq.
        Dim i As Long
        
128     i = 1
        
130     Do While oReq.readyState <> 4

132         DoEvents

134         Sleep 10
136         i = i + 1

138         If i > 1000 Then

140             i = 1
142             WriteLog "***本内容获取超时***" & strURL

144             If iRetry < 3 Then
                    Dim SleepTime As Single
146                 SleepTime = 5 * iRetry
148                 iRetry = iRetry + 1
150                 WriteLog "第 " & iRetry & " 次重试，在" & SleepTime & "秒之后 " & strURL
152                 MySleep SleepTime
154                 GoTo Retry

                Else

156                 WriteLog "3次重试均告失败。。。" & strURL
158                 iRetry = 0

                End If

                Exit Do

            End If

160         DoEvents

        Loop

162     If oReq.readyState = 4 Then
            
                If oReq.Status = 0 Then
                    StartGetHtml = "oReq.Status = 0"
                    Exit Function
                End If
            
'164         If oReq.Status <> 200 Then
'
'
'168                 StartGetHtml = oReq.responseText
'
'            End If
        
188         If oReq.responseText <> "" Then

                Dim strCookies As String
190             strCookies = ""
            
192             strCookies = oReq.getAllResponseHeaders
            
194             If strCookies <> "" Then
            
                    Dim Reg As VBScript_RegExp_55.RegExp
196                 Set Reg = New VBScript_RegExp_55.RegExp
                    Dim Mc As VBScript_RegExp_55.MatchCollection
                    Dim m As VBScript_RegExp_55.Match
198                 Reg.Global = True
200                 Reg.IgnoreCase = True
202                 Reg.MultiLine = False
204                 Reg.Pattern = "Set-Cookie\: (.*?)=(.*?) path=/; expires="
206                 Set Mc = Reg.Execute(strCookies)
                
208                 If Mc.Count > 0 Then

210                     dicCookies.RemoveAll

212                     For Each m In Mc
                        
214                         dicCookies.Add m.SubMatches(0), m.SubMatches(1)
                    
                        Next
                
                    End If

                End If
                If strCharSet = "" Then
                    strCharSet = "GB2312"
                End If
216             StartGetHtml = BytesToBstr(oReq.responseBody, strCharSet)

218             StartGetHtml = convertCRLF(StartGetHtml)

            Else

220             StartGetHtml = ""

            End If

        Else

222         StartGetHtml = ""

        End If

ExitFunc:

224     Set oReq = Nothing
    
        '<EhFooter>
        Exit Function

StartGetHtml_Err:
        Resume Next
        'Err.Raise vbObjectError + 100, "clsXMLHTTPGetHtml.StartGetHtml at line " & Erl, Err.Description
        '</EhFooter>
End Function

Public Sub ClearAllCookies()
    
    dicCookies.RemoveAll

End Sub

Private Sub SetCookies(ByRef obj As MSXML2.XMLHTTP)

100     If dicCookies.Count > 0 Then
        
            Dim v As Variant
            Dim strResult As String
102         strResult = ""
104         For Each v In dicCookies.Keys
            
106             strResult = strResult & v & "=" & dicCookies.Item(v) & " "
            
            Next
108         If strResult <> "" Then
110             obj.SetRequestHeader "Cookie", Left(strResult, Len(strResult) - 2)
112             obj.SetRequestHeader "Cookie", Left(strResult, Len(strResult) - 2)      'One More Time to Fix the Microsoft's Bug
            End If
        End If

End Sub

Private Function BytesToBstr(ByVal body As Variant, _
                             ByVal Cset As String) As String
        '<EhHeader>
        On Error GoTo BytesToBstr_Err
        '</EhHeader>

        If UBound(body) > 0 Then
            Dim objStream As ADODB.Stream
            Set objStream = New ADODB.Stream
100         objStream.Type = 1
110         objStream.Mode = 3
120         objStream.Open
130         objStream.Write body
140         objStream.Position = 0
150         objStream.Type = 2
160         objStream.CharSet = Cset
170         BytesToBstr = objStream.ReadText
180         objStream.Close
190         Set objStream = Nothing
            '<EhFooter>
        Else
            BytesToBstr = ""
        End If

        Exit Function

BytesToBstr_Err:
        Err.Raise vbObjectError + 100, "in modHTTP.BytesToBstr at line " & Erl, Err.Description
        Resume Next
        '</EhFooter>
End Function

Public Property Get URL() As String
    URL = strURL
End Property

Public Property Let URL(ByVal URL As String)
    URL = Replace(URL, "&amp;", "&", 1, -1, vbBinaryCompare)
    strURL = URL
    strMyPostData = ""
End Property

Public Property Get PostData() As String
    PostData = strMyPostData
End Property

Public Property Let PostData(ByVal PostData As String)
    strMyPostData = PostData
End Property

Public Property Let CharSet(ByVal MyCharSet As String)
    strCharSet = MyCharSet
End Property
Public Property Get Cookies() As Scripting.Dictionary
   Set Cookies = dicCookies
End Property

Public Property Let Cookies(ByVal MydicCookies As Scripting.Dictionary)
    Set dicCookies = MydicCookies
End Property

Private Sub Class_Initialize()
    Set dicCookies = New Scripting.Dictionary
    Set dicRequestHeader = New Scripting.Dictionary
End Sub

Public Property Let Referer(ByVal Referer As String)
    strReferer = Replace(Referer, "&amp;", "&", 1, -1, vbBinaryCompare)
End Property
Public Function SetRequestHeader(ByRef strName As String, ByRef strVal As String, Optional ForceOverWrite As Boolean = False) As Boolean

    If Not dicRequestHeader.Exists(strName) Then
    
        dicRequestHeader.Add strName, strVal
        SetRequestHeader = True
    Else
        If ForceOverWrite Then
            dicRequestHeader.Item(strName) = strVal
            SetRequestHeader = True
        Else
            SetRequestHeader = False
        End If
    End If

End Function
Private Sub SetRequestHeaders(ByRef obj As MSXML2.XMLHTTP)

    If dicRequestHeader.Count > 0 Then
        
        Dim v As Variant

        For Each v In dicRequestHeader.Keys
            
            obj.SetRequestHeader CStr(v), dicRequestHeader.Item(v)
            
        Next
      
    End If

End Sub

