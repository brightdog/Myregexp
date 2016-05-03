Attribute VB_Name = "modCRLF"
Option Explicit


Public Function convertCRLF(ByRef value As String) As String
    convertCRLF = Replace(value, vbCr, Chr$(3))
    convertCRLF = Replace(convertCRLF, vbLf, Chr$(4))
End Function

Public Function restoreCRLF(ByRef value As String) As String
    restoreCRLF = Replace(value, Chr$(3), vbCr)
    restoreCRLF = Replace(restoreCRLF, Chr$(4), vbLf)
End Function

