VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTestFTS 
   Caption         =   "FastTextStream Tester"
   ClientHeight    =   3795
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8145
   ForeColor       =   &H00C00000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   3795
   ScaleWidth      =   8145
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtRealLoc 
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   1920
      Width           =   1935
   End
   Begin VB.TextBox txtLocation 
      Height          =   285
      Left            =   2040
      TabIndex        =   23
      Top             =   120
      Width           =   6015
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "&OpenFile"
      Height          =   375
      Left            =   4920
      TabIndex        =   22
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close File"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3360
      TabIndex        =   21
      Top             =   3240
      Width           =   1455
   End
   Begin VB.TextBox txtBufferIndex 
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   2280
      Width           =   1935
   End
   Begin VB.TextBox txtLinesInBuffer 
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   2640
      Width           =   1935
   End
   Begin VB.TextBox txtFileSize 
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   2280
      Width           =   1935
   End
   Begin VB.TextBox txtCurrentLoc 
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   1920
      Width           =   1935
   End
   Begin VB.TextBox txtlastBuffer 
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   1560
      Width           =   6035
   End
   Begin VB.TextBox txtNextBuffer 
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   1200
      Width           =   6035
   End
   Begin VB.TextBox txt1stBuffer 
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   840
      Width           =   6035
   End
   Begin VB.TextBox txtEOF 
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   2640
      Width           =   1935
   End
   Begin VB.TextBox txtReadLine 
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   480
      Width           =   6035
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "&Exit"
      Height          =   375
      Left            =   6480
      TabIndex        =   9
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton cmdSkip 
      Caption         =   "&Skip"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton cmdReadLine 
      Caption         =   "&ReadLine"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3240
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   5280
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label11 
      Caption         =   "Real File Position"
      Height          =   255
      Left            =   4200
      TabIndex        =   26
      Top             =   1920
      Width           =   1650
   End
   Begin VB.Label Label10 
      Caption         =   "File Name"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   120
      Width           =   1650
   End
   Begin VB.Label Label9 
      Caption         =   "Buffer Index"
      Height          =   255
      Left            =   4200
      TabIndex        =   18
      Top             =   2280
      Width           =   1650
   End
   Begin VB.Label Label1 
      Caption         =   "Lines in buffer"
      Height          =   255
      Left            =   4200
      TabIndex        =   17
      Top             =   2640
      Width           =   1650
   End
   Begin VB.Label Label8 
      Caption         =   "Last line in buffer"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   1650
   End
   Begin VB.Label Label7 
      Caption         =   "Last Read Line"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   1650
   End
   Begin VB.Label Label6 
      Caption         =   "First line in buffer"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   1650
   End
   Begin VB.Label Label3 
      Caption         =   "Next line in buffer"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   1650
   End
   Begin VB.Label Label5 
      Caption         =   "AtEndOfStream"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2640
      Width           =   1650
   End
   Begin VB.Label Label4 
      Caption         =   "File size"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2280
      Width           =   1650
   End
   Begin VB.Label Label2 
      Caption         =   "Current File Position"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   1650
   End
   Begin VB.Menu mnuCompareTimes 
      Caption         =   "Compare Times"
   End
   Begin VB.Menu mnuABout 
      Caption         =   "About"
      NegotiatePosition=   3  'Right
   End
End
Attribute VB_Name = "frmTestFTS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Faster Reading of larger files
'Jonathan Orgel, kalonymous, Inc.
Const VSS_ID = "$Header: /FastTextStream/Forms/testFTS.frm 4     1/10/99 11:27a Joni $"
'$NoKeywords: $

Option Explicit


Dim objFTS As Object
Dim CurrentLine As Variant
Const constRed = &HFF&
Const constBlue = &HC00000


Private Sub cmdClose_Click()
    Set objFTS = Nothing
    cmdOpen.Enabled = True 'Already have a file open
    cmdReadLine.Enabled = False
    cmdSkip.Enabled = False
    cmdClose.Enabled = False
    CurrentLine = Null
End Sub

Private Sub cmdOpen_Click()
    With dlgCommonDialog
        .FileName = txtLocation.Text
        .Filter = "Text files|*.txt| " _
            & "All files| *.*"
         .Flags = cdlOFNFileMustExist
         .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        txtLocation.Text = .FileName
    End With
    Set objFTS = New FastTextStream
    If objFTS.OpenTextFileForReading(txtLocation.Text) Then
        cmdOpen.Enabled = False 'Already have a file open
        cmdReadLine.Enabled = True
        cmdSkip.Enabled = True
        cmdClose.Enabled = True
        setFieldvalues
    Else
        MsgBox "Error: Could not open " & txtLocation.Text & ".", _
            vbApplicationModal + vbCritical
    End If
End Sub

Private Sub cmdExit_Click()
    UnloadAllForms
End Sub


Sub setFieldvalues()
    Dim cntrlItem As Control
    
    For Each cntrlItem In frmTestFTS.Controls
        If TypeName(cntrlItem) = "TextBox" And cntrlItem.Name <> "txtLocation" Then
            cntrlItem.ForeColor = constBlue
        End If
    Next
    
    txtLocation.Text = objFTS.Path
    If CurrentLine <> Empty Then
        txtReadLine.Text = CurrentLine
    Else
        txtReadLine.Text = "No line read yet or last read attempt failed"
        txtReadLine.ForeColor = constRed
    End If
    
    txtLinesInBuffer.Text = CStr(objFTS.NumberOfBufferLines)
    txtBufferIndex.Text = objFTS.IndexNextBufferLine
     If objFTS.NumberOfBufferLines > 0 Then
        txt1stBuffer.Text = objFTS.getBufferLine(0)
        txtlastBuffer.Text = objFTS.getBufferLine(objFTS.NumberOfBufferLines)
    Else
        txt1stBuffer.Text = "Zero lines in the buffer"
        txtlastBuffer.Text = "Zero lines in the buffer"
        txt1stBuffer.ForeColor = constRed
        txtlastBuffer.ForeColor = constRed
    End If
    
    txtNextBuffer.Text = objFTS.getBufferLine(objFTS.IndexNextBufferLine)
    txtCurrentLoc.Text = objFTS.CurrentPosition
    txtFileSize.Text = objFTS.Size
    txtEOF.Text = CStr(objFTS.AtEndOfFile)
    txtRealLoc.Text = objFTS.RealFilePosition
    
    Me.Refresh
End Sub

Private Sub cmdReadLine_Click()
    CurrentLine = objFTS.ReadLine()
    setFieldvalues
End Sub

Private Sub cmdSkip_Click()
    Dim strInput As String
    
    strInput = InputBox("How many bytes do you want to skip (use a negative number to go back)")
    If IsNumeric(strInput) Then
        objFTS.Skip CLng(strInput)
        setFieldvalues
    Else
        MsgBox "Please enter a number", vbApplicationModal + vbCritical
    End If
End Sub

Private Sub Form_Load()
        CurrentLine = Null
End Sub

Private Sub mnuAbout_Click()
    MsgBox "FastTextStream Tester" & vbNewLine _
        & "by Jonathan Orgel" & vbNewLine _
        & "Kalonymous, Inc" & vbNewLine _
        & vbNewLine & "VSS: " & VSS_ID
End Sub

Private Sub mnuCompareTimes_Click()
    frmTimeFTS.Show
End Sub
