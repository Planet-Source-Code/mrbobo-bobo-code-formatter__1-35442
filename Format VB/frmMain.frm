VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bobo Code Formatter"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   7905
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   5880
      TabIndex        =   10
      Top             =   3960
      Width           =   1815
   End
   Begin VB.CheckBox ChIndent 
      Caption         =   "Standard Indentation"
      Height          =   255
      Left            =   5880
      TabIndex        =   9
      Top             =   1560
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.CheckBox ChRemoveComments 
      Caption         =   "Remove Comments"
      Height          =   255
      Left            =   5880
      TabIndex        =   8
      Top             =   1080
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.CheckBox ChBlanks 
      Caption         =   "Remove Blank Lines"
      Height          =   375
      Left            =   5880
      TabIndex        =   7
      Top             =   600
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5880
      TabIndex        =   6
      Top             =   2520
      Width           =   1815
   End
   Begin MSComctlLib.ProgressBar PB 
      Height          =   375
      Left            =   5880
      TabIndex        =   5
      Top             =   2040
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdParse 
      Caption         =   "Apply"
      Height          =   375
      Left            =   5880
      TabIndex        =   4
      Top             =   3000
      Width           =   1815
   End
   Begin VB.CheckBox ChBU 
      Caption         =   "Always make BackUp"
      Height          =   255
      Left            =   3240
      TabIndex        =   3
      Top             =   240
      Value           =   1  'Checked
      Width           =   2055
   End
   Begin VB.FileListBox File1 
      Height          =   3795
      Left            =   3000
      Pattern         =   "*.frm;*.mod;*.cls;*.ctl;*.pag;*.dsr"
      TabIndex        =   2
      Top             =   600
      Width           =   2655
   End
   Begin VB.DirListBox Dir1 
      Height          =   3690
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2775
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************
'***************Copyright PSST 2001********************************
'***************Written by MrBobo**********************************
'This code was submitted to Planet Source Code (www.planetsourcecode.com)
'If you downloaded it elsewhere, they stole it and I'll eat them alive

'This little app is a modified version of an AddIn that I use from
'within the IDE - I may upload that in the future. Even if your
'own code is very tidy, when you download other peoples code
'it can sometimes be nearly unreadable because of the formatting
'or an overuse of comments(like this one perhaps). The aim of this
'project is to quickly provide standard indenting, remove blank
'lines and optionally comments from VB source code.
'I hope you find it useful.

Dim CodeString As String, Header As String, TheBigCancel As Boolean
Dim CurFile As String, FullFile As String
Private Sub cmdCancel_Click()
    TheBigCancel = True 'whooh nellie! Bail out please
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdParse_Click()
    'file to parse
    CurFile = IIf(Right(File1.Path, 1) = "\", File1.Path, File1.Path + "\") + File1.List(File1.ListIndex)
    If Not FileExists(CurFile) Then 'error handling
        MsgBox "File not found " & vbCrLf & CurFile
        Exit Sub
    End If
    LoadVB 'open file and split into Header and Code
    If ChBU.Value = 1 Then MakeBU 'backup in case the results are unacceptable
    ParseCode 'OK do it
    FullFile = Header + CodeString 'put it back together after parsing
    If FileExists(CurFile) Then Kill CurFile 'remove old file
    FileSave FullFile, CurFile 'write the new file
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub
Private Sub ParseCode()
    Dim ActualLine As String, TrimmedLine As String, z As Long, mLine As String
    Dim lastTab As Long, EndofDeclarations As Boolean
    PB.Max = UBound(Split(CodeString, vbCrLf)) 'number of lines
    PB.Visible = True 'show progress
    cmdCancel.Enabled = True 'allow bail out
    TheBigCancel = False 'set bail out flag
    For z = 0 To PB.Max 'loop through each line
        DoEvents 'yes please - otherwise we wont be able to press the cancel button
        If TheBigCancel Then Exit Sub 'if we pressed the cancel button bail out now
        ActualLine = Split(CodeString, vbCrLf)(z) 'get a line of code
        'if we're removing comments and the line starts with ' then skip this line
        If ChRemoveComments.Value = 1 And Left(Trim(ActualLine), 1) = "'" Then GoTo nextPlease
        'if we're removing comments then check the line for '
        If ChRemoveComments.Value = 1 Then
            If InStr(Trim(ActualLine), "'") Then 'found a possible comment
                TrimmedLine = CleanComments(ActualLine) 'remove the comment as appropriate
                ActualLine = TrimmedLine 'change ActualLine to new value
            End If
        End If
        'if we're removing blank lines and the line is blank then skip this line
        If ChBlanks.Value = 1 And ChIndent.Value = 0 And Trim(ActualLine) = "" Then GoTo nextPlease
        'If we're not Indenting then go to the next line
        If ChIndent.Value = 0 Then
            mLine = mLine + ActualLine + vbCrLf
            GoTo nextPlease
        End If
        ActualLine = Trim(ActualLine) 'Now we're this far change the ActualLine to a trimmed version
        'Indenting is slightly different for declarations so we need to
        'determine when the first Sub/Function/Property starts
        If Not EndofDeclarations Then
            If Left(ActualLine, 12) = "Private Sub " Then
                EndofDeclarations = True
            ElseIf Left(ActualLine, 17) = "Private Function " Then
                EndofDeclarations = True
            ElseIf Left(ActualLine, 17) = "Private Property " Then
                EndofDeclarations = True
            ElseIf Left(ActualLine, 11) = "Public Sub " Then
                EndofDeclarations = True
            ElseIf Left(ActualLine, 16) = "Public Function " Then
                EndofDeclarations = True
            ElseIf Left(ActualLine, 16) = "Public Property " Then
                EndofDeclarations = True
            ElseIf Left(ActualLine, 4) = "Sub " Then
                EndofDeclarations = True
            ElseIf Left(ActualLine, 9) = "Function " Then
                EndofDeclarations = True
            ElseIf Left(ActualLine, 9) = "Property " Then
                EndofDeclarations = True
            End If
        End If
        'If we're into the Sub/Function/Property bit of code then...
        If EndofDeclarations Then
            'lastTab represents how far to indent the line
            'depending on the content of the line we set lastTab to a new value
            If Left(ActualLine, 3) = "If " Then
                mLine = mLine + String(4 * lastTab, Chr(32)) + ActualLine + vbCrLf
                If Right(ActualLine, 4) = "Then" Then lastTab = lastTab + 1
            ElseIf Left(ActualLine, 7) = "ElseIf " Then
                lastTab = IIf(lastTab < 1, 0, lastTab - 1)
                mLine = mLine + String(4 * lastTab, Chr(32)) + ActualLine + vbCrLf
                If Right(ActualLine, 4) = "Then" Then lastTab = lastTab + 1
            ElseIf ActualLine = "Else" Then
                lastTab = IIf(lastTab < 1, 0, lastTab - 1)
                mLine = mLine + String(4 * lastTab, Chr(32)) + ActualLine + vbCrLf
                lastTab = lastTab + 1
            ElseIf ActualLine = "End If" Then
                lastTab = IIf(lastTab < 1, 0, lastTab - 1)
                mLine = mLine + String(4 * lastTab, Chr(32)) + ActualLine + vbCrLf
            ElseIf Left(ActualLine, 3) = "Do " Then
                mLine = mLine + String(4 * lastTab, Chr(32)) + ActualLine + vbCrLf
                lastTab = lastTab + 1
            ElseIf ActualLine = "Do" Then
                mLine = mLine + String(4 * lastTab, Chr(32)) + ActualLine + vbCrLf
                lastTab = lastTab + 1
            ElseIf Left(ActualLine, 4) = "Loop" Then
                lastTab = IIf(lastTab < 1, 0, lastTab - 1)
                mLine = mLine + String(4 * lastTab, Chr(32)) + ActualLine + vbCrLf
            ElseIf Left(ActualLine, 4) = "For " Then
                mLine = mLine + String(4 * lastTab, Chr(32)) + ActualLine + vbCrLf
                lastTab = lastTab + 1
            ElseIf ActualLine = "Next" Then
                lastTab = IIf(lastTab < 1, 0, lastTab - 1)
                mLine = mLine + String(4 * lastTab, Chr(32)) + ActualLine + vbCrLf
            ElseIf Left(ActualLine, 5) = "Next " Then
                lastTab = IIf(lastTab < 1, 0, lastTab - 1)
                mLine = mLine + String(4 * lastTab, Chr(32)) + ActualLine + vbCrLf
            ElseIf Left(ActualLine, 12) = "Select Case " Then
                mLine = mLine + String(4 * lastTab, Chr(32)) + ActualLine + vbCrLf
                lastTab = lastTab + 2
            ElseIf Left(ActualLine, 5) = "Case " Then
                lastTab = IIf(lastTab < 1, 0, lastTab - 1)
                mLine = mLine + String(4 * lastTab, Chr(32)) + ActualLine + vbCrLf
                lastTab = lastTab + 1
            ElseIf Left(ActualLine, 10) = "End Select" Then
                lastTab = lastTab - 2
                mLine = mLine + String(4 * lastTab, Chr(32)) + ActualLine + vbCrLf
            ElseIf Left(ActualLine, 5) = "With " Then
                mLine = mLine + String(4 * lastTab, Chr(32)) + ActualLine + vbCrLf
                lastTab = lastTab + 1
            ElseIf Left(ActualLine, 8) = "End With" Then
                lastTab = IIf(lastTab < 1, 0, lastTab - 1)
                mLine = mLine + String(4 * lastTab, Chr(32)) + ActualLine + vbCrLf
            ElseIf Left(ActualLine, 12) = "Private Sub " Then
                lastTab = 1
                mLine = mLine + ActualLine + vbCrLf
            ElseIf Left(ActualLine, 17) = "Private Function " Then
                lastTab = 1
                mLine = mLine + ActualLine + vbCrLf
            ElseIf Left(ActualLine, 17) = "Private Property " Then
                lastTab = 1
                mLine = mLine + ActualLine + vbCrLf
            ElseIf Left(ActualLine, 11) = "Public Sub " Then
                lastTab = 1
                mLine = mLine + ActualLine + vbCrLf
            ElseIf Left(ActualLine, 16) = "Public Function " Then
                lastTab = 1
                mLine = mLine + ActualLine + vbCrLf
            ElseIf Left(ActualLine, 16) = "Public Property " Then
                lastTab = 1
                mLine = mLine + ActualLine + vbCrLf
            ElseIf Left(ActualLine, 4) = "Sub " Then
                lastTab = 1
                mLine = mLine + ActualLine + vbCrLf
            ElseIf Left(ActualLine, 9) = "Function " Then
                lastTab = 1
                mLine = mLine + ActualLine + vbCrLf
            ElseIf Left(ActualLine, 9) = "Property " Then
                lastTab = 1
                mLine = mLine + ActualLine + vbCrLf
            ElseIf Left(ActualLine, 8) = "Private " Then
                lastTab = 1
                mLine = mLine + ActualLine + vbCrLf
            ElseIf Left(ActualLine, 7) = "Public " Then
                lastTab = 1
                mLine = mLine + ActualLine + vbCrLf
            ElseIf Left(ActualLine, 7) = "End Sub" Then
                lastTab = 1
                mLine = mLine + ActualLine + vbCrLf
            ElseIf Left(ActualLine, 12) = "End Property" Then
                lastTab = 1
                mLine = mLine + ActualLine + vbCrLf
            ElseIf Left(ActualLine, 12) = "End Function" Then
                lastTab = 1
                mLine = mLine + ActualLine + vbCrLf
            Else
                 mLine = mLine + String(4 * lastTab, Chr(32)) + ActualLine + vbCrLf
            End If
        Else 'Declarations are must be done before Sub/Function/Property bit of code
             'or parsing the Sub/Function/Property bit of code becomes difficult
            If Left(ActualLine, 4) = "#If " Then
                mLine = mLine + String(4 * lastTab, Chr(32)) + ActualLine + vbCrLf
                If Right(ActualLine, 4) = "Then" Then lastTab = lastTab + 1
            ElseIf Left(ActualLine, 8) = "#ElseIf " Then
                lastTab = IIf(lastTab < 1, 0, lastTab - 1)
                mLine = mLine + String(4 * lastTab, Chr(32)) + ActualLine + vbCrLf
                If Right(ActualLine, 4) = "Then" Then lastTab = lastTab + 1
            ElseIf ActualLine = "#Else" Then
                lastTab = IIf(lastTab < 1, 0, lastTab - 1)
                mLine = mLine + String(4 * lastTab, Chr(32)) + ActualLine + vbCrLf
                lastTab = lastTab + 1
            ElseIf ActualLine = "#End If" Then
                lastTab = IIf(lastTab < 1, 0, lastTab - 1)
                mLine = mLine + String(4 * lastTab, Chr(32)) + ActualLine + vbCrLf
            ElseIf Left(ActualLine, 13) = "Private Type " Then
                lastTab = 1
                mLine = mLine + ActualLine + vbCrLf
            ElseIf Left(ActualLine, 12) = "Public Type " Then
                lastTab = 1
                mLine = mLine + ActualLine + vbCrLf
            ElseIf Left(ActualLine, 13) = "Private Enum " Then
                lastTab = 1
                mLine = mLine + ActualLine + vbCrLf
            ElseIf Left(ActualLine, 12) = "Public Enum " Then
                lastTab = 1
                mLine = mLine + ActualLine + vbCrLf
            ElseIf Left(ActualLine, 5) = "Type " Then
                lastTab = 1
                mLine = mLine + ActualLine + vbCrLf
            ElseIf Left(ActualLine, 5) = "Enum " Then
                lastTab = 1
                mLine = mLine + ActualLine + vbCrLf
            ElseIf Left(ActualLine, 8) = "End Type" Then
                lastTab = 0
                mLine = mLine + ActualLine + vbCrLf
            ElseIf Left(ActualLine, 8) = "End Enum" Then
                lastTab = 0
                mLine = mLine + ActualLine + vbCrLf
            Else
                 mLine = mLine + String(4 * lastTab, Chr(32)) + ActualLine + vbCrLf
            End If
        End If
nextPlease:
        PB.Value = z 'display progress
    Next
    CodeString = mLine 'return the parsed string
    cmdCancel.Enabled = False 'disable cancel
    PB.Visible = False 'hide the progress bar
    PB.Value = 0

End Sub
Public Function CleanComments(mLine As String) As String
    Dim cm As Long, qu As Long, z As Long, qCount As Long
    cm = InStr(1, mLine, Chr(39)) 'location of first comment mark(')
    'are there any quotes peceding it that may invalidate it?
    qu = InStrRev(mLine, Chr(34), cm)
    If cm <> 0 Then 'there's definately a comment mark
        If qu <> 0 Then 'there's definately a quote mark
            qCount = 0
            'count the quotes before the comment
            For z = 1 To cm
                If Mid(mLine, z, 1) = Chr(34) Then qCount = qCount + 1
            Next
            If qCount = 0 Or (qCount Mod 2) = 0 Then
                'if there are an even number of quotes before the comment
                'then the comment is valid so use it
                CleanComments = Trim(Left(mLine, cm - 1))
            Else
                'otherwise just return the line in full
                CleanComments = Trim(mLine)
            End If
        Else 'there's no quote before the comment so strip everything after the comment mark
            CleanComments = Trim(Left(mLine, cm - 1))
        End If
    Else
        CleanComments = Trim(mLine)
    End If
End Function

Public Sub LoadVB()
    Dim f As Integer, Searchstr As String
    If LCase(Mid$(CurFile, InStrRev(CurFile, ".") + 1)) = "bas" Then
        Searchstr = "Attribute VB_Name = " 'module headers end like this
    Else
        Searchstr = "Attribute VB_Exposed" 'all the other files end like this
    End If
    f = FreeFile
    Open CurFile For Binary As f 'read the file into a variable
    FullFile = String(LOF(f), Chr$(0))
    Get f, , FullFile
    Close f
    fg = InStr(1, FullFile, Searchstr) 'locate the end of the header
    If fg <> 0 Then 'found it
        fg = InStr(fg + 1, FullFile, vbCrLf) 'move to the end of the header's last line
        If fg <> 0 Then
            CodeString = Right(FullFile, Len(FullFile) - fg - 1) 'here's the actual code
            Header = Left(FullFile, Len(FullFile) - Len(CodeString)) 'here's the header
        End If
    Else 'couldn't find it - use the entire file
        CodeString = FullFile
        Header = ""
    End If

End Sub

Public Sub MakeBU()
    Dim temp As String
    'get a unique filename using the original filename with the "bak" extension
    temp = SafeSave(ChangeExt(CurFile, "bak"))
    'save the entire file as a backup in case things go wrong
    'so all you hard work doesn't evaporate
    FileSave FullFile, temp
End Sub
