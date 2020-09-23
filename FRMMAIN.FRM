VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmmain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MSP's Image Viewer"
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10065
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   10065
   StartUpPosition =   1  'CenterOwner
   Begin MSComCtl2.UpDown udInt 
      Height          =   375
      Left            =   2820
      TabIndex        =   18
      Top             =   4920
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   661
      _Version        =   393216
      Value           =   1000
      Increment       =   100
      Max             =   10000
      Min             =   100
      Enabled         =   0   'False
   End
   Begin VB.CommandButton cmdextview 
      Caption         =   "&External Viewer"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   17
      Top             =   5640
      Width           =   1575
   End
   Begin VB.CheckBox chkcyclic 
      Caption         =   "&Cyclic Slide Show"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   15
      Top             =   4560
      Width           =   2295
   End
   Begin VB.CommandButton cmddeselectall 
      Caption         =   "&DeSelect All"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   14
      Top             =   5640
      Width           =   1575
   End
   Begin VB.CommandButton cmdselectall 
      Caption         =   "Select &All"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   13
      Top             =   5640
      Width           =   1575
   End
   Begin VB.TextBox txttimer 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1800
      TabIndex        =   11
      Text            =   "1000"
      Top             =   4920
      Width           =   1215
   End
   Begin VB.FileListBox File2 
      Height          =   3990
      Left            =   3240
      MultiSelect     =   1  'Simple
      Pattern         =   "*.jpg"
      TabIndex        =   10
      Top             =   360
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Frame framode 
      Caption         =   "Mode"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   6960
      TabIndex        =   7
      Top             =   4200
      Width           =   1695
      Begin VB.OptionButton optmulti 
         Caption         =   "M&ultiple"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton optsingle 
         Caption         =   "S&ingle"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdfull 
      Caption         =   "&Full Screen"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   6
      Top             =   5640
      Width           =   1575
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      TabIndex        =   5
      Top             =   5640
      Width           =   1575
   End
   Begin VB.ComboBox comPattern 
      Height          =   315
      ItemData        =   "frmmain.frx":000C
      Left            =   3240
      List            =   "frmmain.frx":001C
      TabIndex        =   4
      Text            =   "*.jpg"
      Top             =   0
      Width           =   3135
   End
   Begin VB.CheckBox chkpic 
      Caption         =   " Show Preview"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      TabIndex        =   3
      Top             =   3720
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.FileListBox File1 
      Height          =   3990
      Left            =   3240
      Pattern         =   "*.jpg"
      TabIndex        =   2
      Top             =   360
      Width           =   3135
   End
   Begin VB.DirListBox Dir1 
      Height          =   3915
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3135
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3135
   End
   Begin VB.Image imgpic 
      BorderStyle     =   1  'Fixed Single
      Height          =   2655
      Left            =   6600
      Stretch         =   -1  'True
      Top             =   240
      Width           =   3255
   End
   Begin VB.Label lblpreview 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Preview"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1485
      Left            =   7320
      TabIndex        =   16
      Top             =   1200
      Width           =   1665
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Interval (milli Seonds)"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   840
      TabIndex        =   12
      Top             =   4560
      Width           =   2205
   End
   Begin VB.Image imgscreen 
      Height          =   3900
      Left            =   6000
      Picture         =   "frmmain.frx":003C
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   4575
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim s As Integer
Dim path As String
Dim i As Integer
Dim ti As Long

Private Sub chkpic_Click()
    On Error Resume Next
    If chkpic.Value = Unchecked Then
        imgpic.Picture = LoadPicture()
        lblpreview.Visible = True
    ElseIf chkpic.Value = Checked Then
        If File1.ListIndex > 0 Then
            imgpic.Picture = LoadPicture(path)
            lblpreview.Visible = False
        Else
            lblpreview.Visible = True
        End If
    End If
End Sub

Private Sub cmddeselectall_Click()
  On Error Resume Next
  For i = 0 To File2.ListCount - 1
        File2.Selected(i) = False
  Next
  cmdselectall.Enabled = True
  cmddeselectall.Enabled = False
End Sub

Private Sub cmdexit_Click()
    On Error Resume Next
    Unload Me
    Unload frmfullscreen
    Unload frmslideshow
    Unload frmextview
    End
End Sub

Private Sub cmdextview_Click()
   On Error Resume Next
   If optsingle.Value = True Then
   If File1.ListIndex >= 0 Then
        images(0) = path
        Load frmextview
    If frmmain.optsingle.Value = True Then
        frmextview.cmdnext.Visible = False
        frmextview.cmdprevious.Visible = False
        frmextview.cmdslide.Visible = False
        frmextview.cmdviewline2.Visible = False
        tis = 1
    ElseIf frmmain.optmulti.Value = True Then
        frmextview.cmdnext.Visible = True
        frmextview.cmdprevious.Visible = True
        frmextview.cmdslide.Visible = True
        frmextview.cmdviewline2.Visible = True
    End If
        frmextview.Show 1
   Else
        s = MsgBox("No Image Selected For External Viewer", vbInformation, "External Viewer Error")
   End If
   ElseIf optmulti.Value = True Then
       For i = 0 To 9999
            images(i) = " "
        Next
        If Val(txttimer.Text) < 100 Then
            txttimer.Text = "1000"
            s = MsgBox("Minimum Value is 100 Millseconds", vbCritical + vbOKOnly, "Interval Underflow Error")
            Exit Sub
        ElseIf Val(txttimer.Text) > 10000 Then
            txttimer.Text = "1000"
            s = MsgBox("Maximum Value is 10000 Millseconds", vbCritical + vbOKOnly, "Interval Overflow Error")
            Exit Sub
        End If
        tis = 0
            For i = 0 To File2.ListCount - 1
                If File2.Selected(i) = True Then
                    images(tis) = Dir1.path & "\" & File2.List(i)
                    tis = tis + 1
                End If
            Next
        If tis > 0 Then
            Load frmextview
            frmextview.Show 1
        End If
   End If
End Sub

Private Sub cmdfull_Click()
    On Error Resume Next
    If cmdfull.Caption = "&Full Screen" Then
        If File1.ListIndex >= 0 Then
            frmfullscreen.Image1.Picture = LoadPicture(path)
            frmfullscreen.Show 1
        Else
            s = MsgBox("No Image Selected For Full Screen View", vbInformation, "Full Screen View Error")
        End If
    ElseIf cmdfull.Caption = "&Slide Show" Then
        For i = 0 To 9999
            images(i) = " "
        Next
        If File2.ListIndex <> -1 Then
        If IsNumeric(txttimer.Text) Then
            If Val(txttimer.Text) < 100 Then
                s = MsgBox("Minimum Value is 100 Millseconds", vbCritical + vbOKOnly, "Interval Underflow Error")
                txttimer.Text = "1000"
                Exit Sub
            ElseIf Val(txttimer.Text) > 10000 Then
                s = MsgBox("Maximum Value is 10000 Millseconds", vbCritical + vbOKOnly, "Interval Overflow Error")
                txttimer.Text = "1000"
            Exit Sub
            End If
        tis = 0
            For i = 0 To File2.ListCount - 1
                If File2.Selected(i) = True Then
                    images(tis) = Dir1.path & "\" & File2.List(i)
                    tis = tis + 1
                End If
            Next
        If tis >= 0 Then
            frmslideshow.Show 1
        Else
            s = MsgBox("No Image Selected For Slide Show", vbInformation, "External Viewer Error")
        End If
        Else
            s = MsgBox("Interval Cannot Be Non-Numeric", vbCritical + vbOKOnly, "Timer Value Error")
            txttimer.Text = "1000"
        End If
        Else
            s = MsgBox("No Image Selected For Slide Show", vbInformation, "Full Screen View Error")
        End If
    End If
End Sub

Private Sub cmdselectall_Click()
  On Error Resume Next
  For i = 0 To File2.ListCount - 1
        File2.Selected(i) = True
  Next
  cmddeselectall.Enabled = True
  cmdselectall.Enabled = False
End Sub

Private Sub comPattern_Change()
    On Error Resume Next
    File1.Pattern = comPattern.Text
    File2.Pattern = comPattern.Text
End Sub

Private Sub comPattern_Click()
    On Error Resume Next
    File1.Pattern = comPattern.Text
    File2.Pattern = comPattern.Text
End Sub

Private Sub Dir1_Change()
    On Error Resume Next
    File1.path = Dir1.path
    File2.path = Dir1.path
End Sub

Private Sub Drive1_Change()
    On Error GoTo errorhandel
    Dir1.path = Drive1.Drive
errorhandel:
    If Err.Number = 68 Then
        s = MsgBox("Drive " & UCase$(Drive1.Drive) & " Not Ready", vbCritical + vbOKOnly, "MSP Drive Select Error")
        Drive1.Drive = "C:\"
    End If
End Sub

Private Sub File1_Click()
    On Error Resume Next
    If chkpic.Value = Checked Then
        path = Dir1.path & "\" & File1.List(File1.ListIndex)
        imgpic.Picture = LoadPicture(path)
        lblpreview.Visible = False
    ElseIf chkpic.Value = Unchecked Then
        path = Dir1.path & "\" & File1.List(File1.ListIndex)
    End If
End Sub

Private Sub Form_Load()
    Drive1.Drive = "C:\"
End Sub

Private Sub optmulti_Click()
    On Error Resume Next
    cmdfull.Caption = "&Slide Show"
    For i = 0 To File2.ListCount - 1
        File2.Selected(i) = False
    Next
    chkcyclic.Enabled = True
    cmdselectall.Enabled = True
    cmddeselectall.Enabled = False
    File1.Visible = False
    File2.Visible = True
    txttimer.Enabled = True
    udInt.Enabled = True
    Label1.Enabled = True
    lblpreview.Caption = " Preview Not Available In MultiMode"
    chkpic.Enabled = False
    imgpic.Picture = LoadPicture()
    lblpreview.Visible = True
End Sub

Private Sub optsingle_Click()
    On Error Resume Next
    cmdfull.Caption = "&Full Screen"
    chkcyclic.Enabled = False
    cmdselectall.Enabled = False
    cmddeselectall.Enabled = False
    File1.Visible = True
    File2.Visible = False
    lblpreview.Caption = "Preview"
    txttimer.Enabled = False
    udInt.Enabled = False
    Label1.Enabled = False
    chkpic.Enabled = True
    If File1.ListIndex >= 0 Then
        lblpreview.Visible = False
        imgpic.Picture = LoadPicture(path)
    End If
End Sub

Private Sub udInt_Change()
    On Error Resume Next
    txttimer.Text = udInt.Value
End Sub
