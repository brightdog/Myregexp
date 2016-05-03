VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWinHTTPGetHtml"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub Sleep _
                Lib "kernel32" (ByVal dwMilliseconds As Long)
Private strURL As String
Private strMyPostData As String
Private strCharSet As String
Private strReferer As String
Private strLastURL As String
Private iRetry As Integer
Private iRedirectTimes As Integer

Private dicCookies As Scripting.Dictionary
Private dicRequestHeader As Scripting.Dictionary
Private WithEvents oReq As WinHttp.WinHttpRequest
Attribute oReq.VB_VarHelpID = -1
Private bolDataFinished As Boolean
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
        Dim oReq As WinHttp.WinHttpRequest
106     Set oReq = New WinHttp.WinHttpRequest
108     oReq.Option(WinHttpRequestOption_EnableHttp1_1) = False
110     oReq.Option(WinHttpRequestOption_EnableRedirects) = True
112     oReq.Option(WinHttpRequestOption_MaxAutomaticRedirects) = iRedirectTimes
114     oReq.setTimeouts 30000, 30000, 30000, 60000
        ' 第一个数值：解析DNS名字的超时时间
        ' 第二个数值：建立Winsock连接的超时时间
        ' 第三个数值：发送数据的超时时间
        ' 第四个数值：接收response的超时时间
        
116     oReq.Option(WinHttpRequestOption_SslErrorIgnoreFlags) = 13056
        
Retry:

118     'oReq.setProxy 2, "127.0.0.1:8888"
        oReq.setProxy 0
        'HTTPREQUEST_PROXYSETTING_DEFAULT   = 0;
        'HTTPREQUEST_PROXYSETTING_PRECONFIG = 0;
        'HTTPREQUEST_PROXYSETTING_DIRECT    = 1;
        'HTTPREQUEST_PROXYSETTING_PROXY     = 2;
120     bolDataFinished = False
        
122     If strMyPostData = "" Then

124         oReq.Open "GET", strURL, True

126         If Not dicRequestHeader.Exists("User-Agent") Then
128             oReq.SetRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 5.1) AppleWebKit/536.5 (KHTML, like Gecko) Chrome/19.0.1084.46 Safari/536.5"
            End If

130         oReq.SetRequestHeader "Accept", "text/html"
            'oReq.SetRequestHeader "Proxy-Connection", "keep-alive"
132         oReq.SetRequestHeader "Pragma", "no-cache"

134         If strCharSet <> "" Then
136             oReq.SetRequestHeader "Accept-Charset", strCharSet
            End If

138         Call SetCookies(oReq)
140         Call SetRequestHeaders(oReq)

142         If strReferer <> "" Then
144             oReq.SetRequestHeader "Referer", strReferer
            Else

146             If strLastURL <> "" Then
148                 oReq.SetRequestHeader "Referer", strLastURL
                End If
            End If

150         MySleep 0.2
152         oReq.send Null

        Else

154         oReq.Open "POST", strURL, True
156         oReq.SetRequestHeader "Content-Length", Len(strMyPostData)
158         oReq.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
            'oReq.SetRequestHeader "Proxy-Connection", "keep-alive"

160         If Not dicRequestHeader.Exists("User-Agent") Then
162             oReq.SetRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 5.1) AppleWebKit/536.5 (KHTML, like Gecko) Chrome/19.0.1084.46 Safari/536.5"
            End If

164         oReq.SetRequestHeader "Accept", "text/json, text/javascript, text/html"
166         oReq.SetRequestHeader "Pragma", "no-cache"

168         If strCharSet <> "" Then
170             oReq.SetRequestHeader "Accept-Charset", strCharSet
            End If

172         Call SetCookies(oReq)
174         Call SetRequestHeaders(oReq)

176         If strReferer <> "" Then
178             oReq.SetRequestHeader "Referer", strReferer
            Else

180             If strLastURL <> "" Then
182                 oReq.SetRequestHeader "Referer", strLastURL
                End If
            End If

184         oReq.send strMyPostData
            
        End If
        
186     strLastURL = strURL
        'oreq.
        Dim i As Long
        
188     i = 1
        
190     Do While Not bolDataFinished

192         Debug.Print "111"
            
194         MySleep 0.2

196         If oReq.Status = 200 Then
198             Debug.Print "222"
                Exit Do
                
200         ElseIf oReq.Status = 503 Then
202             StartGetHtml = "Status: " & oReq.Status & ", 第 " & iRetry & " 次重试失败！"
204             GoTo ExitFunc
206         ElseIf oReq.Status = 404 Then
            
208             If Len(oReq.responseText) < 50 Then
210                 WriteLog oReq.responseText & strURL & "**" & PostData
212                 StartGetHtml = oReq.responseText
                    Exit Function
                End If
            
            End If

            'Debug.Print TypeName(oReq.Status)
214         DoEvents

216         Sleep 10
218         i = i + 1

220         If i > 600 Then

222             i = 1
224             WriteLog "***本内容获取超时***" & strURL

226             If iRetry < 3 Then
                    Dim SleepTime As Single
228                 SleepTime = 5 * iRetry
230                 iRetry = iRetry + 1
232                 WriteLog "第 " & iRetry & " 次重试，在" & SleepTime & "秒之后 " & strURL
234                 MySleep SleepTime
236                 GoTo Retry

                Else

238                 WriteLog "3次重试均告失败。。。" & strURL
240                 iRetry = 0

                End If

                Exit Do

            End If

242         DoEvents

        Loop

244     If oReq.Status = 200 Then
            
            '            If oReq.Status = 503 Then
            '                StartGetHtml = ""
            '                Exit Function
            '            End If

246         If oReq.statusText <> "OK" Then

WaitRetry:

248             DoEvents
                Dim strLocation As String
250             strLocation = ""
252             strLocation = oReq.getResponseHeader("Location")
                        
254             If strLocation <> "" Then
                    
256                 strURL = GetNewURL(strURL, strLocation)
                    WriteLog strURL
                    Debug.Print strURL
258                 oReq.abort
260                 GoTo Retry '暴力一点了。没办法了。
                    
                End If

262             If iRetry >= 30 Then

264                 StartGetHtml = "Status: " & oReq.Status & ":" & oReq.statusText & ", 第 " & iRetry & " 次重试失败！"
266                 WriteLog StartGetHtml
268                 GoTo ExitFunc

                End If

270             iRetry = iRetry + 1
                'WriteLog "Status: " & oReq.Status & ":" & oReq.statusText & ", 第 " & iRetry & " 次重试，在" & iSleepTime * iRetry / 10 & "秒之后 " & strUrl

272             WriteLog "Wait: " & iRetry & ":" & iSleepTime * iRetry / 10
274             MySleep iSleepTime * iRetry / 10

276             If oReq.statusText <> "OK" Then
278                 Debug.Print "-999:" & iRetry
280                 GoTo WaitRetry
                End If

            End If
            
282         If oReq.statusText <> "OK" Then
284             Debug.Print "-9999"
286             GoTo Retry
            End If
        
288         If oReq.responseText <> "" Then
290             Debug.Print "333"
                Dim strHeaders As String
292             strHeaders = ""
            
294             strHeaders = oReq.getAllResponseHeaders
            
296             If strHeaders <> "" Then
                    Dim arrHeader() As String
                    
298                 arrHeader = Split(strHeaders, vbCrLf, -1, vbBinaryCompare)
                    
300                 If UBound(arrHeader) >= 0 Then
                        
                        Dim iHeader As Integer
                        
                        Dim Reg As VBScript_RegExp_55.RegExp
302                     Set Reg = New VBScript_RegExp_55.RegExp
                        Dim Mc As VBScript_RegExp_55.MatchCollection
                        Dim m As VBScript_RegExp_55.Match
304                     Reg.Global = True
306                     Reg.IgnoreCase = True
308                     Reg.MultiLine = False
310                     Reg.Pattern = "Set-Cookie\:\s*([^=]+=.*?)$"

                        'dicCookies.RemoveAll
312                     For iHeader = 0 To UBound(arrHeader)

314                         Set Mc = Reg.Execute(arrHeader(iHeader))
                
316                         If Mc.Count > 0 Then



318                             For Each m In Mc
                                    Dim arr(1) As String
                        
320                                 arr(0) = Left(m.SubMatches(0), InStr(1, m.SubMatches(0), "=", vbBinaryCompare) - 1)
322                                 arr(1) = Right(m.SubMatches(0), Len(m.SubMatches(0)) - Len(arr(0)) - 1)

324                                 If Not dicCookies.Exists(arr(0)) Then
326                                     dicCookies.Add arr(0), arr(1)
                                    End If
                    
                                Next
                
                            End If

                        Next
                        
328                     Set Reg = Nothing
330                     Set Mc = Nothing
332                     Set m = Nothing
                    End If
                End If
                
334             If strCharSet = "" Then
336                 strCharSet = "GB2312"
                End If

338             Debug.Print "444"
340             iRetry = 0
RetryResponseBody:
342             StartGetHtml = BytesToBstr(oReq.responseBody, strCharSet)
                
344             If StartGetHtml = "" Then
                
346                 If iRetry < 20 Then
348                     WriteLog iRetry
350                     iRetry = iRetry + 1
                        WriteLog iRetry & ":" & "RetryResponseBody"
                        Debug.Print iRetry & ":" & "RetryResponseBody"
352                     MySleep 5
                        DoEvents
354                     GoTo RetryResponseBody

                    End If
                End If
                
356             Debug.Print "555" & "|" & Len(StartGetHtml)
358             StartGetHtml = convertCRLF(StartGetHtml)

            Else
360             Debug.Print "-444"
362             StartGetHtml = oReq.responseText

            End If

        Else
364         Debug.Print "-666"
366         StartGetHtml = oReq.responseText

        End If

ExitFunc:
368     oReq.abort
370     Set oReq = Nothing
    
        '<EhFooter>
        Exit Function

StartGetHtml_Err:

        If Err.Number = -2147483638 Then
            Resume Next
        Else
        
            WriteLog Err.Number & "|" & Err.Description
            
        End If

        'Debug.Print TypeName(oReq.Status) & "*"
        Resume Next
        'Err.Raise vbObjectError + 100, "clsXMLHTTPGetHtml.StartGetHtml at line " & Erl, Err.Description
        '</EhFooter>
End Function


Private Sub SetCookies(ByRef obj As WinHttp.WinHttpRequest)

    If dicCookies.Count > 0 Then
        
        Dim v As Variant
        Dim strResult As String
        strResult = ""
        For Each v In dicCookies.Keys
            
            strResult = strResult & v & "=" & dicCookies.Item(v) & "; "
            
        Next
        If strResult <> "" Then
            obj.SetRequestHeader "Cookie", Left(strResult, Len(strResult) - 2)
            'obj.AddRequestHeader "Cookie", Left(strResult, Len(strResult) - 2)      'One More Time to Fix the Microsoft's Bug
        End If
    End If

End Sub

Private Sub SetRequestHeaders(ByRef obj As WinHttp.WinHttpRequest)

    If dicRequestHeader.Count > 0 Then
        
        Dim v As Variant

        For Each v In dicRequestHeader.Keys
            
            obj.SetRequestHeader CStr(v), dicRequestHeader.Item(v)
            
        Next
      
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
    strReferer = ""
End Property

Public Property Let Referer(ByVal Referer As String)
    strReferer = Replace(Referer, "&amp;", "&", 1, -1, vbBinaryCompare)
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
Public Property Let RedirectTimes(ByVal RedirectTimes As Integer)
    iRedirectTimes = RedirectTimes
End Property
Private Sub Class_Initialize()
    Set dicCookies = New Scripting.Dictionary
    Set dicRequestHeader = New Scripting.Dictionary
    iRedirectTimes = 0  '默认不跳转
End Sub

Public Sub ClearAllCookies()
    
    dicCookies.RemoveAll

End Sub

Public Sub ClearAllRequestHeaders()

    dicRequestHeader.RemoveAll
End Sub

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

Public Function SetRequestCookie(ByRef strName As String, ByRef strVal As String, Optional ForceOverWrite As Boolean = False) As Boolean

    If Not dicCookies.Exists(strName) Then
    
        dicCookies.Add strName, strVal
        SetRequestCookie = True
    Else
        If ForceOverWrite Then
            dicCookies.Item(strName) = strVal
            SetRequestCookie = True
        Else
            SetRequestCookie = False
        End If
    End If

End Function

Private Sub oReq_OnResponseDataAvailable(Data() As Byte)
    bolDataFinished = True
End Sub

Private Sub oReq_OnResponseFinished()
    bolDataFinished = True
End Sub

Private Function GetHeaderLocation(ByRef strHeaders As String) As String

    Dim i As Integer
    Dim strResult As String
    Dim arr() As String
    
    strResult = ""
    arr = Split(strHeaders, vbCrLf, -1, vbBinaryCompare)
    
    For i = 0 To UBound(arr)
    
        If InStr(1, arr(i), ": ", vbBinaryCompare) > 0 Then
        
            Dim arrTmp() As String
            
            arrTmp = Split(arr(i), ": ", -1, vbBinaryCompare)
            
            If UBound(arrTmp) > 0 Then
            
                If arrTmp(0) = "Location" Then
                
                    strResult = Trim(arrTmp(1))
                    Exit For
                End If
            
            Else
            
            End If
        
        End If
    
    Next
    
    GetHeaderLocation = strResult

End Function

Private Function GetNewURL(strURL, strLocation) As String
    
    If Left(LCase(strLocation), 7) = "http://" Or Left(LCase(strLocation), 8) = "https://" Then
        GetNewURL = strLocation
    
    Else
        
        Dim strHost As String
    
        strHost = GetHost(strURL)
    
        '需要区分是否从根路径出发的location
        If Left(strLocation, 1) = "/" Then
            '从根路径出发的！
            GetNewURL = "http://" & strHost & strLocation
    
        Else
    
            GetNewURL = ""
        End If
    End If

End Function

Private Function GetHost(strURL) As String
    
    Dim Reg As VBScript_RegExp_55.RegExp
    Set Reg = New VBScript_RegExp_55.RegExp
    
    Dim Mc As VBScript_RegExp_55.MatchCollection
    
    Reg.Global = False
    Reg.IgnoreCase = True
    Reg.MultiLine = False
    
    Reg.Pattern = "://(.*?)[/\\?#]"
    
    Set Mc = Reg.Execute(strURL)
    
    If Mc.Count > 0 Then
    
        GetHost = Mc.Item(0).SubMatches(0)
    
    Else
        GetHost = ""
    End If
    
    
    Set Reg = Nothing
    Set Mc = Nothing
    
End Function


Private Function GetCurrentPath(strURL) As String
    
    Dim Reg As VBScript_RegExp_55.RegExp
    Set Reg = New VBScript_RegExp_55.RegExp
    
    Dim Mc As VBScript_RegExp_55.MatchCollection
    
    Reg.Global = False
    Reg.IgnoreCase = True
    Reg.MultiLine = False
    
    Reg.Pattern = "://(.*?)[?#]"
    
    Set Mc = Reg.Execute(strURL)
    
    If Mc.Count > 0 Then
    
    Dim arr() As String

    arr = Split(Mc.Item(0).SubMatches(0), "/", -1, vbBinaryCompare)
    
    Dim i As Integer
    Dim Result As String

    For i = 0 To UBound(arr) - 1
    
        Result = Result & arr(i) & "/"
    
    Next

        GetCurrentPath = Result
    
    Else
        GetCurrentPath = ""
    End If
    
    
    Set Reg = Nothing
    Set Mc = Nothing
    
End Function
