VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form waifu2xGUI 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "waifu2x"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Enabled         =   0   'False
   LinkTopic       =   "waifu2x"
   MaxButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check3 
      Caption         =   "Check2"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   2040
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Run"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1680
      TabIndex        =   13
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Browse"
      Height          =   255
      Left            =   3480
      TabIndex        =   12
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      TabIndex        =   11
      Top             =   2400
      Width           =   3135
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComCtl2.UpDown ScaleRatioUpDown 
      Height          =   285
      Left            =   4186
      TabIndex        =   10
      Top             =   1680
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   503
      _Version        =   393216
      Value           =   20
      BuddyControl    =   "ScaleRatio"
      BuddyDispid     =   196613
      OrigLeft        =   4320
      OrigTop         =   1680
      OrigRight       =   4575
      OrigBottom      =   1935
      Max             =   40
      Min             =   10
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown NoiseLevelUpDown 
      Height          =   285
      Left            =   4201
      TabIndex        =   9
      Top             =   1320
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   503
      _Version        =   393216
      Value           =   3
      BuddyControl    =   "NoiseLevel"
      BuddyDispid     =   196614
      OrigLeft        =   4320
      OrigTop         =   1320
      OrigRight       =   4575
      OrigBottom      =   1575
      Max             =   3
      Min             =   1
      Enabled         =   -1  'True
   End
   Begin VB.TextBox cTextBox 
      Height          =   285
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox ScaleRatio 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "2"
      Top             =   1680
      Width           =   2625
   End
   Begin VB.TextBox NoiseLevel 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "3"
      Top             =   1320
      Width           =   2640
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1680
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.Frame Processors 
      Caption         =   "Processors"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.ListBox ProcessorList 
         Height          =   645
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.Label Label3 
      Caption         =   " Use pngcrush to attempt compressing the result"
      Height          =   255
      Left            =   480
      TabIndex        =   15
      Top             =   2040
      Width           =   3495
   End
   Begin VB.Label Label2 
      Caption         =   " Scale ratio:"
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   " Noise level:"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   1320
      Width           =   975
   End
End
Attribute VB_Name = "waifu2xGUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private NoiseBool As Boolean
Private ScaleBool As Boolean
Private Const SpacePattern = "\s\s+"
Private Const Space = " "
Private Const Quote = """"
Private Const Png = "png"
Private Const ClosingParen = ")"
Private Const OpenCl = "(OpenCL )"
Private SpaceRegex As RegExp
Private Cores() As String
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
    ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1

Private Function GetNoiseScaleArg() As String
    Dim mask As Long
    Dim result As String
    mask = 0
    mask = IIf(ScaleBool, mask + 1, mask)
    mask = IIf(NoiseBool, mask + 2, mask)
    Select Case mask
        Case 0
            result = ""
        Case 1
            result = "--scale_ratio " & ScaleRatio.Text & Space & _
                "-m scale" & Space
        Case 2
            result = "--noise_level " & NoiseLevel.Text & Space & _
                "-m noise" & Space
        Case 3
            result = "--scale_ratio " & ScaleRatio.Text & Space & _
                "--noise_level " & NoiseLevel.Text & Space & _
                "-m noise_scale" & Space
    End Select
    GetNoiseScaleArg = result
End Function

Private Sub Check2_Click()
    ScaleBool = Not ScaleBool
    ScaleRatio.Enabled = ScaleBool
    ScaleRatioUpDown.Enabled = ScaleBool
End Sub

Private Sub Check1_Click()
    NoiseBool = Not NoiseBool
    NoiseLevel.Enabled = NoiseBool
    NoiseLevelUpDown.Enabled = NoiseBool
End Sub

Private Sub Command1_Click()
    CommonDialog1.Filter = "All files (*.*)|*.*"
    CommonDialog1.DefaultExt = ""
    CommonDialog1.DialogTitle = "Select Image File"
    CommonDialog1.ShowOpen
    Text1.Text = CommonDialog1.FileName
End Sub

Private Sub Command2_Click()
    Dim arguments As String
    arguments = "waifu2x-converter-cpp.exe " & _
        "--processor " & ProcessorList.ListIndex & Space & _
        "-j " & Cores(ProcessorList.ListIndex) & Space & _
        GetNoiseScaleArg() & "-i " & Quote & Text1.Text & Quote
    Call U.Exec(arguments, App.path, False)
    If Not Check3.Enabled Then
        GoTo SkipCrush
    End If
    Dim output As String
    Dim sFilename As String
    Dim path As String
    path = App.path & "\"
    sFilename = Dir$(path & "*.*")
    Do While LenB(sFilename) <> 0
        If StrComp(LCase(Right(sFilename, 3)), Png) = 0 Then
            If LenB(output) <> 0 Then
                Dim date1 As Date
                Dim date2 As Date
                date1 = FileDateTime(path & output)
                date2 = FileDateTime(path & sFilename)

                If date2 > date1 Then
                    output = sFilename
                End If
            Else
                output = sFilename
            End If
        End If
        sFilename = Dir$()
    Loop
    If LenB(output) <> 0 Then
        Call U.Exec("pngcrush -ow " & Quote & output & Quote, App.path, False)
    End If
SkipCrush:
    Unload Me
    End
End Sub

Private Sub Form_Load()
    Set SpaceRegex = New RegExp
    SpaceRegex.Pattern = SpacePattern
    SpaceRegex.Global = True
    NoiseBool = True
    ScaleBool = True

    Dim i As Integer
    Dim iCore As Integer
    Dim Processors() As String
    Dim Processor As String
    Dim LastParenPos As Long
    Dim strLen As Long

    Call U.ExecAndCapture("waifu2x-converter-cpp.exe --list-processor", cTextBox, App.path)
    If LenB(cTextBox.Text$) = 0 Then
        MsgBox Quote & "waifu2x-converter-cpp.exe" & Quote & _
            " was not found in the PATH or there are no compatible processors to use with waifu2x!", _
            vbCritical, "Missing executable"
        Unload Me
        End
    End If

    Processors = Split(cTextBox.Text, IIf(InStrB(cTextBox.Text$, vbNewLine$) <> 0, vbNewLine, vbLf))
    iCore = 0
    For i = LBound(Processors) To UBound(Processors)
        Processor = Trim(Processors(i))
        If LenB(Processor$) <> 0 Then
            Processor = SpaceRegex.Replace(Processor, Space)
            ReDim Preserve Cores(iCore)
            Cores(iCore) = Right$(Processor, 1)
            iCore = iCore + 1
            LastParenPos = InStrRev(Processor, ClosingParen) + 1
            strLen = Len(Processor)
            Processor = Mid$(Processor, 4, strLen - (strLen - LastParenPos) - 4)
            ProcessorList.AddItem (Processor)
            If InStrB(Processor$, OpenCl$) <> 0 Then
                ProcessorList.ListIndex = iCore - 1
            End If
        End If
    Next

    Set SpaceRegex = Nothing

    waifu2xGUI.Enabled = True
    Call SetWindowPos(waifu2xGUI.hwnd, -1, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
End Sub

Private Sub NoiseLevelUpDown_Change()
    NoiseLevel.Text = Str$(NoiseLevelUpDown.Value)
End Sub

Private Sub ScaleRatioUpDown_Change()
    ScaleRatio.Text = Str$(ScaleRatioUpDown.Value / 10)
End Sub

Private Sub Text1_Change()
    If LenB(Text1.Text) <> 0 Then
        Command2.Enabled = True
    End If
End Sub
