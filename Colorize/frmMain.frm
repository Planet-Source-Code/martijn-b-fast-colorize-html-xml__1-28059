VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Fast Colorize"
   ClientHeight    =   1605
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2400
   LinkTopic       =   "Form1"
   ScaleHeight     =   1605
   ScaleWidth      =   2400
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkView 
      Caption         =   "Auto-view File"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Colorize Big File (500 kb)"
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   2175
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Colorize Normal File (50 kb)"
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdTest_Click(Index As Integer)
   Dim lngCount         As Long
   Dim strFileData      As String
   Dim lngFreeFile      As Long

   On Error Resume Next
   Kill App.Path & "\col.rtf"
   Kill App.Path & "\test.htm"
   On Error GoTo 0

   Select Case Index
      Case 0
         '***  colorize a small file - copy it
         FileCopy App.Path & "\demo.htm", App.Path & "\test.htm"
      Case 1
         '***  colorize a big file - copy data 10 times
         strFileData = fcnGetFile(App.Path & "\demo.htm")
         lngFreeFile = FreeFile

         Open App.Path & "\test.htm" For Append As #lngFreeFile

         For lngCount = 1 To 10
            Print #lngFreeFile, strFileData
         Next

         Close #lngFreeFile
   End Select

   fcnTestCol
End Sub

Sub fcnTestCol()

   Dim strRTF           As String
   Dim sinStart         As Single
   Dim sinEnd           As Single

   sinStart = Timer
   If fcnColorize(App.Path & "\test.htm", strRTF) = True Then
      sinEnd = Timer
      MsgBox "colorized in " & sinEnd - sinStart & " s. output bytes: " & Len(strRTF)
      fcnPutFile App.Path & "\col.rtf", strRTF

      '***  launch WordPad
      If chkView Then Shell "write.exe " & App.Path & "\col.rtf", vbNormalFocus
   Else
      MsgBox "Colorize Failed...", vbExclamation, ""
   End If

End Sub
