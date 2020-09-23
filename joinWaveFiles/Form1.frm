VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Form1"
   ClientHeight    =   10665
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   15240
   DrawMode        =   0  'Blackness
   ForeColor       =   &H00800000&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10665
   ScaleWidth      =   15240
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Join Files"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   3735
      Left            =   6360
      TabIndex        =   28
      Top             =   1920
      Width           =   4215
      Begin VB.PictureBox Picture13 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   240
         Picture         =   "Form1.frx":058A
         ScaleHeight     =   195
         ScaleWidth      =   2235
         TabIndex        =   32
         Top             =   480
         Width           =   2295
         Begin VB.PictureBox Picture15 
            Height          =   255
            Left            =   3600
            ScaleHeight     =   255
            ScaleWidth      =   15
            TabIndex        =   33
            Top             =   0
            Width           =   15
         End
         Begin VB.Label Label4 
            BackColor       =   &H00E0E0E0&
            Caption         =   "File Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   177
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   360
            TabIndex        =   34
            Top             =   0
            Width           =   975
         End
      End
      Begin VB.CommandButton cmdMakeNewFile 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Join Files"
         DisabledPicture =   "Form1.frx":0F8C
         Height          =   1095
         Left            =   2760
         Picture         =   "Form1.frx":440E
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   840
         Width           =   1215
      End
      Begin VB.ListBox lstNewFiles 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000005&
         Height          =   2790
         Left            =   240
         TabIndex        =   29
         Top             =   720
         Width           =   2295
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Commands List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   4335
      Left            =   10680
      TabIndex        =   24
      Top             =   5760
      Width           =   4455
      Begin VB.PictureBox Picture14 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   600
         Picture         =   "Form1.frx":7890
         ScaleHeight     =   195
         ScaleWidth      =   3195
         TabIndex        =   26
         Top             =   480
         Width           =   3255
         Begin VB.Label Label3 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Command Lines"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   177
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   480
            TabIndex        =   27
            Top             =   0
            Width           =   1695
         End
      End
      Begin VB.ListBox lstCommands 
         BackColor       =   &H00E0E0E0&
         Height          =   2985
         Left            =   600
         TabIndex        =   25
         Top             =   720
         Width           =   3255
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Files to Join List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   3735
      Left            =   0
      TabIndex        =   12
      Top             =   1920
      Width           =   6255
      Begin VB.PictureBox Picture9 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   360
         Picture         =   "Form1.frx":8292
         ScaleHeight     =   195
         ScaleWidth      =   1515
         TabIndex        =   22
         Top             =   360
         Width           =   1575
         Begin VB.Label Label2 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Details"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   177
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   360
            TabIndex        =   23
            Top             =   0
            Width           =   615
         End
      End
      Begin VB.ListBox lstFiles 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00FFFFFF&
         Height          =   2595
         Left            =   480
         TabIndex        =   21
         Top             =   600
         Width           =   4815
      End
      Begin VB.CommandButton cmdDelFile 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   3240
         Width           =   4695
      End
      Begin VB.PictureBox Picture8 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   1920
         Picture         =   "Form1.frx":881C
         ScaleHeight     =   195
         ScaleWidth      =   3315
         TabIndex        =   17
         Top             =   360
         Width           =   3375
         Begin VB.PictureBox Picture11 
            Height          =   255
            Left            =   3600
            ScaleHeight     =   255
            ScaleWidth      =   15
            TabIndex        =   18
            Top             =   0
            Width           =   15
         End
         Begin VB.Label Label1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "File Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   177
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   360
            TabIndex        =   19
            Top             =   0
            Width           =   975
         End
      End
      Begin VB.PictureBox Picture10 
         BackColor       =   &H00E0E0E0&
         Height          =   3255
         Left            =   120
         ScaleHeight     =   3195
         ScaleWidth      =   195
         TabIndex        =   16
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox Picture12 
         BackColor       =   &H00E0E0E0&
         Height          =   3255
         Left            =   5400
         ScaleHeight     =   3195
         ScaleWidth      =   675
         TabIndex        =   13
         Top             =   360
         Width           =   735
         Begin VB.CommandButton cmdUp 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   177
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            Picture         =   "Form1.frx":921E
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   960
            Width           =   495
         End
         Begin VB.CommandButton cmdDown 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   177
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            Picture         =   "Form1.frx":9AE8
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   1680
            Width           =   495
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Wave File Graph"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   4335
      Left            =   0
      TabIndex        =   2
      Top             =   5760
      Width           =   10575
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00E0E0E0&
         Height          =   3000
         Left            =   120
         ScaleHeight     =   2940
         ScaleWidth      =   10035
         TabIndex        =   5
         Top             =   720
         Width           =   10100
         Begin VB.PictureBox Picture5 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000001&
            Height          =   1240
            Left            =   0
            ScaleHeight     =   1215
            ScaleWidth      =   10005
            TabIndex        =   9
            Top             =   1520
            Visible         =   0   'False
            Width           =   10040
            Begin VB.PictureBox Picture6 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   1240
               Left            =   1800
               ScaleHeight     =   1245
               ScaleWidth      =   1815
               TabIndex        =   10
               Top             =   0
               Visible         =   0   'False
               Width           =   1815
               Begin VB.PictureBox Picture7 
                  Appearance      =   0  'Flat
                  AutoRedraw      =   -1  'True
                  BackColor       =   &H80000001&
                  ForeColor       =   &H80000005&
                  Height          =   1240
                  Left            =   480
                  ScaleHeight     =   1215
                  ScaleWidth      =   585
                  TabIndex        =   11
                  Top             =   0
                  Width           =   615
                  Begin VB.Line Line4 
                     BorderColor     =   &H80000005&
                     Visible         =   0   'False
                     X1              =   240
                     X2              =   240
                     Y1              =   0
                     Y2              =   1200
                  End
               End
            End
            Begin VB.Line Line2 
               Visible         =   0   'False
               X1              =   600
               X2              =   600
               Y1              =   0
               Y2              =   1200
            End
         End
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H8000000D&
            Height          =   2480
            Left            =   0
            ScaleHeight     =   2445
            ScaleWidth      =   10005
            TabIndex        =   6
            Top             =   240
            Width           =   10040
            Begin VB.PictureBox Picture3 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   2480
               Left            =   1200
               ScaleHeight     =   2475
               ScaleWidth      =   2175
               TabIndex        =   7
               Top             =   0
               Visible         =   0   'False
               Width           =   2175
               Begin VB.PictureBox Picture4 
                  Appearance      =   0  'Flat
                  AutoRedraw      =   -1  'True
                  BackColor       =   &H8000000D&
                  ForeColor       =   &H80000009&
                  Height          =   2480
                  Left            =   600
                  ScaleHeight     =   2445
                  ScaleWidth      =   705
                  TabIndex        =   8
                  Top             =   -30
                  Width           =   735
                  Begin VB.Line Line3 
                     BorderColor     =   &H80000005&
                     Visible         =   0   'False
                     X1              =   360
                     X2              =   360
                     Y1              =   0
                     Y2              =   2400
                  End
               End
            End
            Begin VB.Line Line1 
               Visible         =   0   'False
               X1              =   600
               X2              =   600
               Y1              =   0
               Y2              =   2400
            End
         End
      End
      Begin VB.CommandButton cmdPlay 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Play"
         Height          =   375
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   3840
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "stop"
         Height          =   375
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   3840
         Width           =   615
      End
      Begin VB.Label lblGraphFileName 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   2880
         TabIndex        =   31
         Top             =   360
         Width           =   5175
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdOpenFile 
      BackColor       =   &H00E0E0E0&
      Caption         =   "open"
      Height          =   735
      Left            =   1680
      Picture         =   "Form1.frx":A3B2
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      Width           =   975
   End
   Begin MCI.MMControl MMControl1 
      Height          =   330
      Left            =   4080
      TabIndex        =   0
      Top             =   10320
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   582
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Menu menuFile 
      Caption         =   "Wave File"
      Begin VB.Menu menuFileOpenWave 
         Caption         =   "Open Wave File"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub fillLstFiles()


End Sub



Private Sub cmdDelFile_Click()
    Dim tmpDelFile As String ' put the name of the file to delete
    Dim i, j As Integer
    
    If lstFiles.ListCount = 0 Then ' list is empty

        lstCommands.AddItem "Error Message:"
        lstCommands.AddItem "No file to delete"
        Exit Sub
    End If
    
    If MsgBox("hgjhg", vbYesNo) = vbYes Then
        'temp for copying data
        ReDim tmpwave(numOfFiles - 1) As waveFileInfo
        
        tmpDelFile = wavefiles(curFile).fileName
        
        j = 0
        For i = 0 To UBound(wavefiles)
            If i <> curFile Then
                tmpwave(j) = wavefiles(i)
                j = j + 1
            End If
        Next
        numOfFiles = numOfFiles - 1
        lstFiles.Clear
  
        ReDim wavefiles(numOfFiles - 1)
        For i = 0 To numOfFiles - 1
            curFile = i
            wavefiles(i) = tmpwave(i)
            lstFiles.AddItem getWaveFileTitle
        Next
        
  
        lstCommands.AddItem "File Deleted:"
        lstCommands.AddItem "            " & tmpDelFile
  End If
End Sub

Private Sub cmdMakeNewFile_Click()
curFile = 0
    If checkAllWaveFiles Then
           ' open wave file using common dialog
        CommonDialog1.Filter = "Wave Files (*.wav)|*.wav"
        CommonDialog1.CancelError = True
        On Error GoTo Errhandler
          CommonDialog1.Action = 2
        If CommonDialog1.fileName <> "" Then
            SaveWave CommonDialog1.fileName
            lstNewFiles.AddItem CommonDialog1.fileName
            lstCommands.AddItem "File Saved:"
            lstCommands.AddItem "            " & CommonDialog1.fileName
            Exit Sub
        End If
    End If
Errhandler:
    Exit Sub
End Sub


Private Sub Command1_Click()
MMControl1.Command = "stop"
End Sub

Private Sub Form_Load()
    Me.WindowState = 2
    MMControl1.Visible = False
    numOfFiles = 0
    ReDim wavefiles(numOfFiles)
End Sub

Private Sub cmdOpenFile_Click()

End Sub

Private Sub cmdPlay_Click()
    MMControl1.Command = "close"
    MMControl1.Notify = False
    MMControl1.Wait = True
    MMControl1.Shareable = False ' Set properties needed by MCI to open.
    MMControl1.DeviceType = "WaveAudio"
    'MMControl1.TimeFormat = mciFormatMilliseconds
    Me.lblGraphFileName = Replace(Me.lblGraphFileName, "\", "/")
    MMControl1.fileName = Me.lblGraphFileName
    MMControl1.Command = "Open"
    MMControl1.Command = "Play"



End Sub


Private Sub cmdUp_Click()
    If lstFiles.ListCount = 0 Then Exit Sub
    ReDim temp(lstFiles.ListCount) As String
    ReDim tmpwave(lstFiles.ListCount) As waveFileInfo
    i = lstFiles.ListIndex
    If i = 0 Then
        Exit Sub
    End If
    
    For u = 0 To lstFiles.ListCount - 1
        If u = lstFiles.ListIndex - 1 Then
            temp(lstFiles.ListIndex - 1) = lstFiles.List(lstFiles.ListIndex)
            temp(lstFiles.ListIndex) = lstFiles.List(lstFiles.ListIndex - 1)
            
            tmpwave(lstFiles.ListIndex - 1) = wavefiles(lstFiles.ListIndex)
            tmpwave(lstFiles.ListIndex) = wavefiles(lstFiles.ListIndex - 1)
                    
            u = u + 1
        Else
            temp(u) = lstFiles.List(u)
            tmpwave(u) = wavefiles(u)
        End If
    Next
    lstFiles.Clear
    For j = 0 To UBound(temp) - 1
        lstFiles.AddItem temp(j)
        wavefiles(j) = tmpwave(j)
    Next
    lstCommands.AddItem "File Moved Up:"
    lstCommands.AddItem "            " & wavefiles(curFile).fileName
    
    
    lstFiles.ListIndex = i - 1
End Sub
Private Sub cmdDown_Click()
    If lstFiles.ListCount = 0 Then Exit Sub
    ReDim temp(lstFiles.ListCount) As String
    ReDim tmpwave(lstFiles.ListCount) As waveFileInfo
    i = lstFiles.ListIndex
    If i = lstFiles.ListCount - 1 Then
        Exit Sub
    End If
    For u = 0 To lstFiles.ListCount - 1
        If u = lstFiles.ListIndex Then
            temp(lstFiles.ListIndex) = lstFiles.List(lstFiles.ListIndex + 1)
            temp(lstFiles.ListIndex + 1) = lstFiles.List(lstFiles.ListIndex)
            
            tmpwave(lstFiles.ListIndex) = wavefiles(lstFiles.ListIndex + 1)
            tmpwave(lstFiles.ListIndex + 1) = wavefiles(lstFiles.ListIndex)
            
            u = u + 1
        Else
            temp(u) = lstFiles.List(u)
            tmpwave(u) = wavefiles(u)
        End If
    Next
    lstFiles.Clear
    For j = 0 To UBound(temp) - 1
        lstFiles.AddItem temp(j)
        wavefiles(j) = tmpwave(j)
    
    Next
    lstCommands.AddItem "File Moved Down:"
    lstCommands.AddItem "            " & wavefiles(curFile).fileName

    lstFiles.ListIndex = i + 1

End Sub

Private Sub lstFiles_Click()
   curFile = lstFiles.ListIndex
   Text1 = curFile
End Sub

Private Sub lstFiles_DblClick()
 
    GraphWave wavefiles(curFile).channel, wavefiles(curFile).lenData, _
            wavefiles(curFile).fileName, wavefiles(curFile).numOfBytes, _
            wavefiles(curFile).inData
    lblGraphFileName.Caption = wavefiles(curFile).fileName
End Sub

Private Sub lstNewFiles_DblClick()
    If lstNewFiles.ListCount = 0 Then Exit Sub
        
    newWaveFile.fileName = lstNewFiles.Text
    getWaveFileData newWaveFile
    GraphWave newWaveFile.channel, newWaveFile.lenData, _
            newWaveFile.fileName, newWaveFile.numOfBytes, _
            newWaveFile.inData
    lblGraphFileName.Caption = newWaveFile.fileName
End Sub

Private Sub menuFileOpenWave_Click()
        ' open wave file using common dialog
    CommonDialog1.Filter = "Wave Files (*.wav)|*.wav"
    CommonDialog1.CancelError = True
    On Error GoTo Errhandler
    CommonDialog1.ShowOpen
    If CommonDialog1.fileName <> "" Then
        
        numOfFiles = numOfFiles + 1 ' add another file to nomOfFiles
        reDimWaveFilesArray     ' re dim file array
        wavefiles(numOfFiles - 1).fileName = CommonDialog1.fileName ' put file name in type

        curFile = numOfFiles - 1    ' current file to work on
        getWaveFileData wavefiles(curFile)             ' put all data to type
        
        lstCommands.ForeColor = &H80000007
        lstCommands.AddItem "File Added:"
        lstCommands.AddItem "            " & wavefiles(curfilr).fileName
        
        lstFiles.AddItem getWaveFileTitle
        lstFiles.ListIndex = lstFiles.ListCount - 1 ' select last file entered
        lstFiles_DblClick ' get data and

        Exit Sub
    End If
Errhandler:
    Exit Sub
End Sub
