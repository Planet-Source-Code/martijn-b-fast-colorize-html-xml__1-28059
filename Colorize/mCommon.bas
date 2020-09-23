Attribute VB_Name = "mCommon"
Option Explicit

Function fcnPutFile(strFileName As String, strData As String) As String
   Dim lngFile          As Long
   Dim strFile          As String

   lngFile = FreeFile
   Open strFileName For Output As lngFile
   Print #lngFile, strData;
   Close lngFile
   fcnPutFile = strFileName
End Function

Function fcnGetFile(strFileName As String) As String
   Dim lngFile          As Long
   Dim strFile          As String

   If Dir$(strFileName) <> vbNullString Then
      lngFile = FreeFile
      Open strFileName For Binary Access Read As lngFile
      strFile = Space$(LOF(lngFile))
      Get #lngFile, 1, strFile
      Close lngFile
      fcnGetFile = strFile
   End If
End Function

Function fcnEscapeRTF(strText As String) As String
   '***  escape the rtf special chars
   strText = Replace(strText, "\", "\\")
   strText = Replace(strText, "{", "\{")
   strText = Replace(strText, "}", "\}")
   fcnEscapeRTF = Replace(strText, Chr$(13), "\par" & Chr$(13))
End Function
