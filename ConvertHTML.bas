Attribute VB_Name = "ConvertHTML"
Option Explicit


Public Function ConvertHTML(ByVal Content As String)
    Content = restoreCRLF(Content)
    Content = Replace(Content, vbTab, " ", 1, -1, vbBinaryCompare)
    Content = Replace(Content, "&nbsp;", " ", 1, -1, vbBinaryCompare)
    Content = Replace(Content, "'", "`", 1, -1, vbBinaryCompare)
    Content = Replace(Content, "&lt;", "<", 1, -1, vbBinaryCompare)
    Content = Replace(Content, "&gt;", ">", 1, -1, vbBinaryCompare)
    Content = Replace(Content, Chr$(10), "", 1, -1, vbBinaryCompare)
    Content = Replace(Content, Chr$(9), "", 1, -1, vbBinaryCompare)
    Content = Replace(Content, Chr$(13), "", 1, -1, vbBinaryCompare)
    Content = Replace(Content, "<br>", vbCrLf, 1, -1, vbBinaryCompare)
    Content = Replace(Content, "<BR>", vbCrLf, 1, -1, vbBinaryCompare)
    Content = Replace(Content, "<br />", vbCrLf, 1, -1, vbBinaryCompare)
    Content = Replace(Content, vbCrLf & vbCrLf, vbCrLf, 1, -1, vbBinaryCompare)

    '以上的顺序是有讲究的，不可乱动！
    Dim i As Integer

    For i = 0 To 4
        Content = Replace(Content, "  ", "", 1, -1, vbBinaryCompare)
    Next

    Dim regTmp As VBScript_RegExp_55.RegExp

    Set regTmp = New VBScript_RegExp_55.RegExp
    regTmp.Global = True
    regTmp.MultiLine = True
    regTmp.IgnoreCase = True
    '======================= add by brightdog 去除页面中的干扰码
    regTmp.Pattern = "(<span[^>]*?display\s*?:\s*?none[^>]*?>[\w\W]*?<\/span>)"            '<span style="display:none">/ u6 i* t4 {1 Z. f5 m$ B. H" P1 u</span><br />
    Content = regTmp.Replace(Content, "")
    regTmp.Pattern = "<font([^>]+)(0px|0pt)+([^>]*)>([\w\W]*?)<\/font>"            '<span style="display:none">/ u6 i* t4 {1 Z. f5 m$ B. H" P1 u</span><br />
    Content = regTmp.Replace(Content, "")
    regTmp.Pattern = "<span[^>]*?font\s*?-\s*?size\s*?:\s*(0px|0pt)[^>]*?>([\w\W]*?)<\/span>"            '<span style="display:none">/ u6 i* t4 {1 Z. f5 m$ B. H" P1 u</span><br />
    Content = regTmp.Replace(Content, "")
    regTmp.Pattern = "<script[^>]*?>([\w\W]*?)<\/script>"            '<span style="display:none">/ u6 i* t4 {1 Z. f5 m$ B. H" P1 u</span><br />
    Content = regTmp.Replace(Content, "")
    '=======================
    regTmp.Pattern = "(width\s*>\s*\d+)"            '没办法，好多论坛中的图片带有JS脚本，只能这样试试看了。
    Content = regTmp.Replace(Content, "")
    regTmp.Pattern = "(height\s*>\s*\d+)"            '没办法，好多论坛中的图片带有JS脚本，只能这样试试看了。
    Content = regTmp.Replace(Content, "")
     'regTmp.Pattern = "(<em>.*?</em>)"            '没办法，好多论坛中的图片带有JS脚本，只能这样试试看了。
    'Content = regTmp.Replace(Content, "")
    regTmp.Pattern = "(<(""[^""]*""|'[^']*'|[^'"">])*>)"            '网上说匹配html标记的。不知道灵不灵了。
    Content = regTmp.Replace(Content, "")
    regTmp.Pattern = "(<.*?[^>]>)"            '没办法，好多论坛中的图片带有JS脚本，只能这样试试看了。
    Content = regTmp.Replace(Content, "")
    regTmp.Pattern = "(<.*?>)"
    Content = regTmp.Replace(Content, "")
    ConvertHTML = Trim$(Content)
    Set regTmp = Nothing
End Function


