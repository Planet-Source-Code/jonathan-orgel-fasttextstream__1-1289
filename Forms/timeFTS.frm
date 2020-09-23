VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTimeFTS 
   Caption         =   "Read File Timing Tests"
   ClientHeight    =   5655
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   8160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   8160
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCompare 
      Caption         =   "&Compare Times"
      Height          =   375
      Left            =   5040
      TabIndex        =   25
      Top             =   5160
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   3600
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   6600
      TabIndex        =   0
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "File to Test"
      Height          =   855
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8055
      Begin VB.TextBox txtLocation 
         Height          =   285
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   5895
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "&Browse"
         Height          =   375
         Left            =   6480
         TabIndex        =   2
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame frmTests 
      Caption         =   "Tests"
      Height          =   3975
      Left            =   0
      TabIndex        =   4
      Top             =   1080
      Width           =   8055
      Begin VB.TextBox txtmyskip 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   3240
         Visible         =   0   'False
         Width           =   3695
      End
      Begin VB.CheckBox chkMySkip 
         Caption         =   "Time skipping bytes using FastTextStream skip"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   3240
         Value           =   1  'Checked
         Width           =   3975
      End
      Begin VB.TextBox txtMyReadLine 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   3600
         Visible         =   0   'False
         Width           =   3695
      End
      Begin VB.CheckBox chkMyReadLine 
         Caption         =   "Time reading lines using FastTextStream readline"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   3600
         Value           =   1  'Checked
         Width           =   3975
      End
      Begin VB.CheckBox chkGetRandom 
         Caption         =   "Time reading records using the Get (random,234b)"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   2880
         Value           =   1  'Checked
         Width           =   3975
      End
      Begin VB.TextBox txtGetrandom 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   2880
         Visible         =   0   'False
         Width           =   3695
      End
      Begin VB.CheckBox chkGetBig 
         Caption         =   "Time reading bytes using the Get (binary, 6300b)"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   2520
         Value           =   1  'Checked
         Width           =   3975
      End
      Begin VB.TextBox txtGetBig 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   2520
         Visible         =   0   'False
         Width           =   3695
      End
      Begin VB.TextBox txtGetSmall 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   2160
         Visible         =   0   'False
         Width           =   3695
      End
      Begin VB.TextBox txtSeek 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   1800
         Visible         =   0   'False
         Width           =   3695
      End
      Begin VB.TextBox txtSkip 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1440
         Visible         =   0   'False
         Width           =   3695
      End
      Begin VB.TextBox txtSkipLine 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1080
         Visible         =   0   'False
         Width           =   3695
      End
      Begin VB.TextBox txtRead 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   720
         Visible         =   0   'False
         Width           =   3695
      End
      Begin VB.TextBox txtReadLine 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   360
         Visible         =   0   'False
         Width           =   3695
      End
      Begin VB.CheckBox chkGetSMall 
         Caption         =   "Time reading bytes using the Get (binary, 1000b)"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   2160
         Value           =   1  'Checked
         Width           =   3735
      End
      Begin VB.CheckBox chkSeek 
         Caption         =   "Time skipping bytes using the seek statement"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1800
         Value           =   1  'Checked
         Width           =   3735
      End
      Begin VB.CheckBox chkSkip 
         Caption         =   "Time skipping characters using the .Skip method"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1440
         Value           =   1  'Checked
         Width           =   3855
      End
      Begin VB.CheckBox chkSkipLine 
         Caption         =   "Time skipping lines using .SkipLine method"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1080
         Value           =   1  'Checked
         Width           =   3735
      End
      Begin VB.CheckBox chkRead 
         Caption         =   "Time reading characters using the .Read method"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Value           =   1  'Checked
         Width           =   3855
      End
      Begin VB.CheckBox chkReadLine 
         Caption         =   "Time reading lines using .ReadLine method"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Value           =   1  'Checked
         Width           =   3615
      End
   End
   Begin VB.Menu mnuTest 
      Caption         =   "Test FastTextStream"
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "About"
      NegotiatePosition=   3  'Right
   End
End
Attribute VB_Name = "frmTimeFTS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const VSS_ID = "$Header: /FastTextStream/Forms/timeFTS.frm 9     1/10/99 11:27a Joni $"
'$NoKeywords: $


Private Const constLargeNumber = 2000000000
Private Const RecordSize = 234
Private Type RecBuffer
    Contents(RecordSize) As Byte
End Type
Private Const BigBufferSize = 63000
Private Type BigBuffer
    Contents(BigBufferSize) As Byte
End Type
Private Const SmallBufferSize = 1000
Private Type SmallBuffer
    Contents(SmallBufferSize) As Byte
End Type
Private MySmallBuf As SmallBuffer 'Declared here becasue of 32K stack limit
Private MyBigBuf As BigBuffer 'Declared here becasue of 32K stack limit
Private MyRecBuf As RecBuffer


Private Sub cmdBrowse_Click()
    With dlgCommonDialog
        .Filename = txtLocation.Text
        .Filter = "Text files|*.txt| " _
            & "All files| *.*"
        .Flags = cdlOFNFileMustExist
        .ShowOpen
        If Len(.Filename) = 0 Then
            Exit Sub
        End If
        txtLocation.Text = .Filename
    End With
End Sub

Private Sub cmdExit_Click()
    UnloadAllForms
End Sub


Private Sub cmdCompare_Click()
    Dim LinesToRead As Long
    Dim LinesToSkip As Long
    Dim BytesToSkip As Long
    Dim strMessage As String
    Dim ElapsedSeconds As Single
    Dim cntrlItem As Control
    
    'Keep adding new tests. Don't forget one of the labels
    For Each cntrlItem In Me.Controls
        If TypeName(cntrlItem) = "TextBox" And cntrlItem.Name <> "txtLocation" Then
            cntrlItem.Visible = False
        End If
    Next

    Me.Refresh
    
    If Not FileExists(txtLocation.Text) Then
        txtLocation.SetFocus
        txtLocation.SelStart = 0
        txtLocation.SelLength = Len(txtLocation.Text)

        MsgBox "Error:" & vbNewLine & "The file """ _
            & txtLocation.Text & """ does not exist." & vbNewLine, _
            vbApplicationModal + vbCritical
        Exit Sub
    End If
    
    If chkSkip Then
        BytesToSkip = constLargeNumber
        ElapsedSeconds = doSkipBytesTest(txtLocation.Text, BytesToSkip)
        txtSkip.Text = Format(BytesToSkip, "###,###,###") & " bytes in " _
            & Format(ElapsedSeconds, "Standard") & " seconds"
        txtSkip.Visible = True
        Me.Refresh
    End If
    
    If chkSkipLine Then
        LinesToSkip = constLargeNumber
        LinesToRead = 0
        ElapsedSeconds = doSkipLineTest(txtLocation.Text, LinesToSkip, LinesToRead)
        txtSkipLine.Text = Format(LinesToSkip, "###,###,###") & " lines in " _
            & Format(ElapsedSeconds, "Standard") & " seconds"
        txtSkipLine.Visible = True
        Me.Refresh
    End If

    If chkMyReadLine Then
        LinesToSkip = 0
        LinesToRead = constLargeNumber
        ElapsedSeconds = doMyReadLineTest(txtLocation.Text, LinesToRead)
        txtMyReadLine.Text = Format(LinesToRead, "###,###,###") & " lines in " _
            & Format(ElapsedSeconds, "Standard") & " seconds"
        txtMyReadLine.Visible = True
        Me.Refresh
    End If

    If chkReadLine Then
        LinesToSkip = 0
        LinesToRead = constLargeNumber
        ElapsedSeconds = doSkipLineTest(txtLocation.Text, LinesToSkip, LinesToRead)
        txtReadLine.Text = Format(LinesToRead, "###,###,###") & " lines in " _
            & Format(ElapsedSeconds, "Standard") & " seconds"
        txtReadLine.Visible = True
        Me.Refresh
    End If

    If chkSeek Then
        BytesToSkip = constLargeNumber
        ElapsedSeconds = doSeekTest(txtLocation.Text, BytesToSkip)
        txtSeek.Text = Format(BytesToSkip, "###,###,###") & " bytes in " _
            & Format(ElapsedSeconds, "Standard") & " seconds"
        txtSeek.Visible = True
        Me.Refresh
    End If

    If chkGetBig Then
        BytesToSkip = constLargeNumber
        ElapsedSeconds = doGetTestBig(txtLocation.Text, BytesToSkip)
        txtGetBig.Text = Format(BytesToSkip, "###,###,###") & " bytes in " _
            & Format(ElapsedSeconds, "Standard") & " seconds"
        txtGetBig.Visible = True
        Me.Refresh
    End If
    
    If chkGetSMall Then
        BytesToSkip = constLargeNumber
        ElapsedSeconds = doGetTestSmall(txtLocation.Text, BytesToSkip)
        txtGetSmall.Text = Format(BytesToSkip, "###,###,###") & " bytes in " _
            & Format(ElapsedSeconds, "Standard") & " seconds"
        txtGetSmall.Visible = True
        Me.Refresh
    End If
    
    If chkGetRandom Then
        BytesToSkip = constLargeNumber
        ElapsedSeconds = doGetTestRandom(txtLocation.Text, BytesToSkip)
        txtGetrandom.Text = Format(BytesToSkip, "###,###,###") & " bytes in " _
            & Format(ElapsedSeconds, "Standard") & " seconds"
        txtGetrandom.Visible = True
        Me.Refresh
    End If
    
    
    If chkRead Then
        BytesToSkip = constLargeNumber
        ElapsedSeconds = doReadTest(txtLocation.Text, BytesToSkip)
        txtRead.Text = Format(BytesToSkip, "###,###,###") & " bytes in " _
            & Format(ElapsedSeconds, "Standard") & " seconds"
        txtRead.Visible = True
        Me.Refresh
    End If
    
    If chkMySkip Then
        BytesToSkip = constLargeNumber
        ElapsedSeconds = doMySkipBytesTest(txtLocation.Text, BytesToSkip)
        txtmyskip.Text = Format(BytesToSkip, "###,###,###") & " bytes in " _
            & Format(ElapsedSeconds, "Standard") & " seconds"
        txtmyskip.Visible = True
        Me.Refresh
    End If

End Sub


Function doSkipLineTest( _
    Location As String, LinesToSkip As Long, LinesToRead As Long) As Single

    Dim objFso As Object
    Dim objTextStream As Object
    Dim StartTime As Single
    Dim LinesSkipped As Long
    Dim LinesRead As Long

    'This may take a while
    Me.MousePointer = vbHourglass
    
    Set objFso = CreateObject("Scripting.FileSystemObject")
    Set objTextStream = objFso.opentextfile(Location, 1)
    
    StartTime = Timer
    LinesSkipped = 0
    Do While LinesSkipped < LinesToSkip And Not objTextStream.atEndOfstream
        objTextStream.SkipLine
        LinesSkipped = LinesSkipped + 1
    Loop
    LinesToSkip = LinesSkipped
    
    Do While LinesRead < LinesToRead And Not objTextStream.atEndOfstream
        objTextStream.ReadLine
        LinesRead = LinesRead + 1
    Loop
    doSkipLineTest = Timer - StartTime
    LinesToRead = LinesRead
    LinesToSkip = LinesSkipped
   
    objTextStream.Close
    Set objFso = Nothing
    Set objTextStream = Nothing
    Me.MousePointer = vbDefault
End Function

Function doMyReadLineTest(Location As String, LinesToRead As Long) As Single
    Dim objFastTextStream As Object
    Dim StartTime As Single
    Dim LinesRead As Long
 
    'This may take a while
    Me.MousePointer = vbHourglass

    Set objFastTextStream = New FastTextStream
    If Not objFastTextStream.OpenTextFileForReading(Location) Then
        doMyReadLineTest = 0
        LinesToRead = 0
        Exit Function
    End If
    
    LinesRead = 0
    StartTime = Timer
    Do While Not IsEmpty(objFastTextStream.ReadLine)
        LinesToRead = LinesToRead - 1
        LinesRead = LinesRead + 1
    Loop
    doMyReadLineTest = Timer - StartTime
    LinesToRead = LinesRead
   
    Set objFastTextStream = Nothing
    Me.MousePointer = vbDefault
End Function

Function doSkipBytesTest(Location As String, BytesToSkip As Long) As Single
    Dim objFso As Object
    Dim objFile As Object
    Dim objTextStream As Object
    Dim StartTime As Single
    Dim FileSize As Long
 
    'This may take a while
    Me.MousePointer = vbHourglass

    Set objFso = CreateObject("Scripting.FileSystemObject")
    Set objFile = objFso.GetFile(txtLocation.Text)
    Set objTextStream = objFso.opentextfile(txtLocation.Text, 1)
    
    FileSize = objFile.Size
    If FileSize < BytesToSkip Then BytesToSkip = FileSize - 1
    
    StartTime = Timer
    objTextStream.Skip BytesToSkip
    doSkipBytesTest = Timer - StartTime
    
    Set objTextStream = Nothing
    Set objFile = Nothing
    Set objFso = Nothing

    Me.MousePointer = vbDefault
End Function

Function doMySkipBytesTest(Location As String, BytesToSkip As Long) As Single
    Dim objFastTextStream As Object
    Dim StartTime As Single
    Dim FileSize As Long
 
    'This may take a while
    Me.MousePointer = vbHourglass

    Set objFastTextStream = New FastTextStream
    If Not objFastTextStream.OpenTextFileForReading(Location) Then
        Exit Function
        doMySkipBytesTest = 0
        BytesToSkip = 0
    End If
    
    FileSize = objFastTextStream.Size
    If FileSize < BytesToSkip Then BytesToSkip = FileSize - 1
    
    StartTime = Timer
    objFastTextStream.Skip BytesToSkip
    doMySkipBytesTest = Timer - StartTime
    
    Set objFastTextStream = Nothing
    
    Me.MousePointer = vbDefault
End Function

Function doReadTest(Location As String, BytesToSkip As Long) As Single
    Dim objFso As Object
    Dim objFile As Object
    Dim objTextStream As Object
    Dim StartTime As Single
    Dim FileSize As Long
    Dim Steps As Long
    Dim Rest As Long
 
    'This may take a while
    Me.MousePointer = vbHourglass

    Set objFso = CreateObject("Scripting.FileSystemObject")
    Set objFile = objFso.GetFile(txtLocation.Text)
    Set objTextStream = objFso.opentextfile(txtLocation.Text, 1)
    
    FileSize = objFile.Size
    If FileSize <= BytesToSkip Then BytesToSkip = FileSize - 1
    
    Steps = BytesToSkip / BigBufferSize
    Rest = BytesToSkip - Steps * BigBufferSize
    StartTime = Timer
    Do While Steps
        objTextStream.Read BigBufferSize
        Steps = Steps - 1
    Loop
    If Rest > 0 Then objTextStream.Read Rest
    doReadTest = Timer - StartTime
    
    Me.MousePointer = vbDefault
End Function

Function doSeekTest(Location As String, BytesToSkip As Long) As Single
    Dim StartTime As Single
    Dim FileSize As Long
    
    'This may take a while
    Me.MousePointer = vbHourglass

    Open Location For Input As #1   ' Open file for input.
    FileSize = LOF(1)               ' Get size of file in bytes.
    If FileSize < BytesToSkip Then BytesToSkip = FileSize - 1
    
    StartTime = Timer
    Seek #1, BytesToSkip
    doSeekTest = Timer - StartTime
    
    Me.MousePointer = vbDefault
    Close #1
End Function


Function doGetTestSmall(Location As String, BytesToSkip As Long) As Single
    Dim StartTime As Single
    Dim FileSize As Long
    Dim Steps As Long   'Number of whole reads
    
    'This may take a while
    Me.MousePointer = vbHourglass

    Open Location For Binary As #1   ' Open file for input.
    FileSize = LOF(1)               ' Get size of file in bytes.
    If FileSize <= BytesToSkip Then BytesToSkip = FileSize - 1
    
    Steps = BytesToSkip / SmallBufferSize
    If BytesToSkip > SmallBufferSize * Steps Then Steps = Steps + 1
    StartTime = Timer
    Do While Steps > 0
        Get #1, , MySmallBuf
        Steps = Steps - 1
    Loop
    doGetTestSmall = Timer - StartTime
    
    Me.MousePointer = vbDefault
    Close #1
End Function


Function doGetTestBig(Location As String, BytesToSkip As Long) As Single
    Dim StartTime As Single
    Dim FileSize As Long
    Dim Steps As Integer   'Number of whole reads
    
    'This may take a while
    Me.MousePointer = vbHourglass

    Open Location For Binary As #1   ' Open file for input.
    FileSize = LOF(1)               ' Get size of file in bytes.
    If FileSize <= BytesToSkip Then BytesToSkip = FileSize - 1
    
    Steps = BytesToSkip / BigBufferSize
    If BytesToSkip > BigBufferSize * Steps Then Steps = Steps + 1
    StartTime = Timer
    Do While Steps > 0
        Get #1, , MyBigBuf
        Steps = Steps - 1
    Loop
    doGetTestBig = Timer - StartTime
    
    Me.MousePointer = vbDefault
    Close #1
End Function


Function doGetTestRandom(Location As String, BytesToRead As Long) As Single
    Dim StartTime As Single
    Dim FileSize As Long
    Dim Steps As Long   'Number of whole reads
    
    'This may take a while
    Me.MousePointer = vbHourglass

    Open Location For Random As #1 Len = Len(MyRecBuf)
    FileSize = LOF(1)               ' Get size of file in bytes.
    If FileSize <= BytesToRead Then BytesToRead = FileSize - 1
    
    Steps = BytesToRead / RecordSize
    If BytesToRead > RecordSize * Steps Then Steps = Steps + 1
    StartTime = Timer
    Do While Steps
        Get #1, , MyRecBuf
        Steps = Steps - 1
    Loop
    doGetTestRandom = Timer - StartTime
    
    Me.MousePointer = vbDefault
    Close #1
End Function

Private Sub mnuAbout_Click()
    MsgBox "FastTextStream Times Tester" & vbNewLine _
        & "by Jonathan Orgel" & vbNewLine _
        & "Kalonymous, Inc" & vbNewLine _
        & vbNewLine & "VSS: " & VSS_ID
        
End Sub

Private Sub mnuTest_Click()
    frmTestFTS.Show
End Sub
