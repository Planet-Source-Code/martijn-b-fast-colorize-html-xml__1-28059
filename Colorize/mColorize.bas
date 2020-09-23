Attribute VB_Name = "mColorize"
Option Explicit
Option Compare Text

Public arrColorize      As Variant
Private lngColorBlockSize As Long

Function fcnColorize(strFileName As String, Optional ByRef strFileData As String) As Boolean
   fcnColorize = False
   Dim strOverLap       As String
   Dim lngFile          As Long
   Dim lngLen           As Long
   Dim lngLOF           As Long
   Dim lngCounter       As Long
   Dim lngNoBlocks      As Long
   Dim lngExtra         As Long
   Dim lngMain          As Long
   Dim strBlockIn       As String
   Dim lngDataPos       As Long
   Dim lngCount         As Long
   Dim strTemp          As String
   Dim blnDoString      As Boolean

   Dim strFileDataHeader As String

   '***  the file read buffer length = string block lenght
   '***  change this number for optimal performance
   '***  values below 200 will slow down (may block => block overhead hight)
   '***  values above 3000 will slow down (long blocks will slow down string search / replace)

   lngColorBlockSize = 1400

   '***  read the colors
   arrColorize = fcnGetColorizeArray()

   '***  build the rtf header
   strFileDataHeader = "{\rtf1\ansi\deff0\deflang2057{\fonttbl{\f0\fmodern\fprq1\fcharset0 Fixedsys;}}"
   strFileDataHeader = strFileDataHeader & "{\colortbl\red0\green0\blue0;"
   '***  add the rtf colors to the header
   For lngCount = LBound(arrColorize, 2) To UBound(arrColorize, 2)
      strFileDataHeader = strFileDataHeader & "\red" & CStr(arrColorize(7, lngCount)) & "\green" & CStr(arrColorize(8, lngCount)) & "\blue" & CStr(arrColorize(9, lngCount)) & ";"
      '***  add spaces to the color defs
      arrColorize(2, lngCount) = arrColorize(2, lngCount) & Chr$(32&)
      arrColorize(3, lngCount) = arrColorize(3, lngCount) & Chr$(32&)
   Next lngCount
   '***  add some extra tabs to the header
   strFileDataHeader = strFileDataHeader & "}\deflang2057\pard\tx120\tx240\tx360\tx480\tx600\tx720\tx840\plain\f0\fs20\cf0 "

   '***  open the file for binary access
   lngFile = FreeFile
   Open strFileName For Binary As lngFile
   lngLOF = LOF(lngFile)

   lngNoBlocks = lngLOF \ lngColorBlockSize
   lngExtra = lngLOF Mod lngColorBlockSize
   lngMain = lngNoBlocks * lngColorBlockSize

   '***  reserve string space for the file
   strFileData = Space$((lngLOF + Len(strFileDataHeader)) * 4)
   Mid$(strFileData, 1, Len(strFileDataHeader)) = strFileDataHeader
   lngDataPos = Len(strFileDataHeader) + 1

   '***  this is the main loop
   For lngCounter = 1 To lngMain Step lngColorBlockSize
      strTemp = Space$(lngColorBlockSize)
      Get #lngFile, lngCounter, strTemp

      '***  escape the rtf special chars
      strTemp = fcnEscapeRTF(strTemp)

      If LenB(strOverLap) = 0 Then
         strBlockIn = strTemp
      Else
         strBlockIn = Space$(Len(strTemp) + Len(strOverLap))
         Mid$(strBlockIn, 1) = strOverLap
         Mid$(strBlockIn, 1 + Len(strOverLap)) = strTemp
         strOverLap = vbNullString
      End If

      fcnColorizeBlock strBlockIn, 0, strOverLap

      '***  store the filedata in strFileData
      Mid$(strFileData, lngDataPos) = strBlockIn
      lngDataPos = lngDataPos + Len(strBlockIn)
   Next
   '*** end of the main loop

   '***  get the last block of the file
   If lngExtra <> 0 Then
      If Not blnDoString Then
         strBlockIn = Space$(lngExtra)
         Get #lngFile, lngCounter, strBlockIn
      End If

      '***  escape the rtf special chars
      strBlockIn = strOverLap & fcnEscapeRTF(strBlockIn)
      strOverLap = vbNullString
      fcnColorizeBlock strBlockIn, 0, strOverLap

      '***  store the filedata in strFileData (include strOverLap)
      Mid$(strFileData, lngDataPos) = strBlockIn & strOverLap
      lngDataPos = lngDataPos + Len(strBlockIn & strOverLap)
   End If

   Close lngFile

   strFileData = Left$(strFileData, lngDataPos)
   Mid$(strFileData, Len(strFileData)) = "}"

   fcnColorize = True
End Function

Function fcnColorizeBlock(strBlock As String, Optional lngState As Long, Optional strOverLap As String, Optional blnRoot = True) As Boolean
   '***  this function replaces the text in the html file with rtf color codes.
   
   Dim strColStart      As String
   Dim strColEnd        As String
   Dim lngCounter       As Long
   Dim lngLen           As Long
   Dim lngPosS          As Long
   Dim lngPosE          As Long
   Dim lngFoundS        As Long
   Dim lngFoundE        As Long
   Dim strLeft          As String
   Dim strMid           As String
   Dim strRight         As String
   Dim strTemp          As String

   lngLen = UBound(arrColorize, 2)

   If lngState > lngLen Or LenB(strBlock) = 0 Then Exit Function
   lngPosS = 1
   lngCounter = lngState

   strColStart = arrColorize(0, lngCounter)
   If Val(strColStart) <> 0 Then
      strColStart = Chr$(Val(strColStart))
   End If

   lngFoundS = InStr(lngPosS, strBlock, strColStart, arrColorize(6, lngCounter))

   If lngFoundS <> 0 And Not (blnRoot = True And arrColorize(5, lngCounter) = False) Then
      '***  found a start

      '***  process left string
      strLeft = Left$(strBlock, lngFoundS - 1)
      lngState = lngCounter + 1
      fcnColorizeBlock strLeft, lngState, strOverLap, blnRoot

      If LenB(strOverLap) = 0 Then

         '***  process mid
         strColEnd = arrColorize(1, lngCounter)

         lngPosE = lngFoundS + 1

         lngFoundE = InStr(lngPosE, strBlock, strColEnd, arrColorize(6, lngCounter))

         If lngFoundE <> 0 Then
            strMid = Mid$(strBlock, lngFoundS, lngFoundE - lngFoundS + Len(strColEnd))
            If arrColorize(4, lngCounter) Then
               lngState = lngCounter + 1
               fcnColorizeBlock strMid, lngState, strOverLap, False
            End If

            strTemp = Space$(10 + Len(strMid))
            Mid$(strTemp, 1) = arrColorize(2, lngCounter)
            Mid$(strTemp, 6) = strMid
            Mid$(strTemp, 6 + Len(strMid)) = arrColorize(3, lngCounter)
            strMid = strTemp

            strRight = Mid$(strBlock, lngFoundE + Len(strColEnd))
            lngState = lngCounter
            fcnColorizeBlock strRight, lngState, strOverLap, blnRoot

            strBlock = strLeft & strMid & strRight
         ElseIf blnRoot Then
            strOverLap = Mid$(strBlock, lngFoundS)
            lngState = lngCounter
            strBlock = strLeft 'left$(strBlock, lngFoundS - 1)
         End If

      Else

         strOverLap = vbNullString
         fcnColorizeBlock strBlock, lngState, strOverLap, blnRoot

      End If

   Else
      '***  nothing found search next delimiter
      lngState = lngCounter + 1
      fcnColorizeBlock strBlock, lngState, strOverLap, blnRoot

   End If

   fcnColorizeBlock = lngState
End Function

Function fcnGetColorizeArray() As Variant
   '***  the array in this function is customizable
   '***  it contains a number of start - end definitions and colors
   '***  items are replaced ordered by index
   
   'start     start text
   'End       end text
   'startrtf  color code for this item
   'endrtf    color code after this item
   'fill      true if the inside of this item must be colored
   'root      true if this item cannot be inside other items. for example a remark "<!-- ... -->"
   'compare   compare mode used to find start and end in the text. vbTextCompare or vbBinaryCompare
   'red       single byte color value for this item
   'green     single byte color value for this item
   'blue      single byte color value for this item

   '***  new definitions can be made to colorize other file types than html
   '***  feel free to change/extend the array
   '***  also change the redim!

   Dim arrCol           As Variant

   ReDim arrCol(9, 6)

   arrCol(0, 0) = "<!--"
   arrCol(1, 0) = "-->"
   arrCol(2, 0) = "\cf1"
   arrCol(3, 0) = "\cf0"
   arrCol(4, 0) = "False"
   arrCol(5, 0) = "True"
   arrCol(6, 0) = "0"
   arrCol(7, 0) = "255"
   arrCol(8, 0) = "25"
   arrCol(9, 0) = "0"

   arrCol(0, 1) = "<script"
   arrCol(1, 1) = "</script>"
   arrCol(2, 1) = "\cf2"
   arrCol(3, 1) = "\cf0"
   arrCol(4, 1) = "False"
   arrCol(5, 1) = "True"
   arrCol(6, 1) = "1"
   arrCol(7, 1) = "40"
   arrCol(8, 1) = "40"
   arrCol(9, 1) = "160"

   arrCol(0, 2) = "<%"
   arrCol(1, 2) = "%>"
   arrCol(2, 2) = "\cf3"
   arrCol(3, 2) = "\cf0"
   arrCol(4, 2) = "False"
   arrCol(5, 2) = "True"
   arrCol(6, 2) = "0"
   arrCol(7, 2) = "160"
   arrCol(8, 2) = "160"
   arrCol(9, 2) = "25"

   arrCol(0, 3) = "<!"
   arrCol(1, 3) = ">"
   arrCol(2, 3) = "\cf4"
   arrCol(3, 3) = "\cf0"
   arrCol(4, 3) = "False"
   arrCol(5, 3) = "True"
   arrCol(6, 3) = "0"
   arrCol(7, 3) = "160"
   arrCol(8, 3) = "160"
   arrCol(9, 3) = "25"

   arrCol(0, 4) = "<"
   arrCol(1, 4) = ">"
   arrCol(2, 4) = "\cf5"
   arrCol(3, 4) = "\cf0"
   arrCol(4, 4) = "True"
   arrCol(5, 4) = "True"
   arrCol(6, 4) = "0"
   arrCol(7, 4) = "20"
   arrCol(8, 4) = "20"
   arrCol(9, 4) = "255"

   arrCol(0, 5) = """"
   arrCol(1, 5) = """"
   arrCol(2, 5) = "\cf6"
   arrCol(3, 5) = "\cf5"
   arrCol(4, 5) = "False"
   arrCol(5, 5) = "False"
   arrCol(6, 5) = "0"
   arrCol(7, 5) = "120"
   arrCol(8, 5) = "120"
   arrCol(9, 5) = "120"

   arrCol(0, 6) = "'"
   arrCol(1, 6) = "'"
   arrCol(2, 6) = "\cf7"
   arrCol(3, 6) = "\cf5"
   arrCol(4, 6) = "False"
   arrCol(5, 6) = "False"
   arrCol(6, 6) = "0"
   arrCol(7, 6) = "120"
   arrCol(8, 6) = "120"
   arrCol(9, 6) = "120"

   fcnGetColorizeArray = arrCol
End Function
