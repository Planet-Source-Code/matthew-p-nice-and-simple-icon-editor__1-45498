VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmIcon 
   Caption         =   "Icon Edit 2003"
   ClientHeight    =   5700
   ClientLeft      =   4200
   ClientTop       =   3090
   ClientWidth     =   7185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   7185
   Begin VB.CommandButton cmdFill 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Fill"
      Height          =   255
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   3960
      Width           =   495
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4815
      Left            =   1920
      ScaleHeight     =   321
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   321
      TabIndex        =   19
      Top             =   480
      Width           =   4815
      Begin VB.PictureBox Picture6 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   600
         Left            =   0
         ScaleHeight     =   40
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   40
         TabIndex        =   24
         Top             =   0
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.PictureBox Picture4 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   600
         Left            =   0
         ScaleHeight     =   40
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   40
         TabIndex        =   23
         Top             =   0
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.PictureBox Picture10 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   600
         Left            =   0
         ScaleHeight     =   40
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   40
         TabIndex        =   20
         Top             =   0
         Visible         =   0   'False
         Width           =   600
      End
      Begin MSComctlLib.ImageList imlToolBar 
         Left            =   0
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   15
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   32
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":0544
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":0A88
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":0FCC
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":1320
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":1644
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":1968
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":1CBC
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":2200
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":2744
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":2A68
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":2D8C
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":30B0
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":35F4
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":3B38
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":407C
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":45C0
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":4B04
               Key             =   ""
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":5048
               Key             =   ""
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":558C
               Key             =   ""
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":5AD0
               Key             =   ""
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":6014
               Key             =   ""
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":6558
               Key             =   ""
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":6A9C
               Key             =   ""
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":6DC0
               Key             =   ""
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":7304
               Key             =   ""
            EndProperty
            BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":7848
               Key             =   ""
            EndProperty
            BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":7D8C
               Key             =   ""
            EndProperty
            BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":82D0
               Key             =   ""
            EndProperty
            BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":8814
               Key             =   ""
            EndProperty
            BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":8D58
               Key             =   ""
            EndProperty
            BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":929C
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5280
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4560
      Top             =   4680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   5655
      Left            =   1560
      TabIndex        =   21
      Top             =   0
      Width           =   5535
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00AB8F8D&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5655
      Left            =   0
      ScaleHeight     =   5655
      ScaleWidth      =   1515
      TabIndex        =   0
      Top             =   0
      Width           =   1515
      Begin VB.Frame Frame2 
         BackColor       =   &H00AB8F8D&
         Caption         =   "Brush size"
         ForeColor       =   &H00FFFFFF&
         Height          =   1095
         Left            =   120
         TabIndex        =   26
         Top             =   4440
         Width           =   1215
         Begin VB.OptionButton Option1 
            BackColor       =   &H00AB8F8D&
            Caption         =   "large"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   29
            Top             =   720
            Width           =   975
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00AB8F8D&
            Caption         =   "medium"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   28
            Top             =   480
            Width           =   975
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00AB8F8D&
            Caption         =   "small"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   480
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   1
         Top             =   120
         Width           =   480
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   480
         ScaleHeight     =   15
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   29
         TabIndex        =   2
         ToolTipText     =   "Colore corrente"
         Top             =   3600
         Width           =   465
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00AB8F8D&
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   720
         Width           =   975
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   120
         X2              =   1320
         Y1              =   4320
         Y2              =   4320
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         X1              =   120
         X2              =   1320
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FFFFFF&
         X1              =   120
         X2              =   1320
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00E0E0E0&
         Height          =   465
         Left            =   480
         Top             =   120
         Width           =   465
      End
      Begin VB.Label lblColour 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   18
         Top             =   1320
         Width           =   225
      End
      Begin VB.Label lblColour 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   17
         Top             =   1320
         Width           =   225
      End
      Begin VB.Label lblColour 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   2
         Left            =   960
         TabIndex        =   16
         Top             =   1320
         Width           =   225
      End
      Begin VB.Label lblColour 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   15
         Top             =   1680
         Width           =   225
      End
      Begin VB.Label lblColour 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   4
         Left            =   600
         TabIndex        =   14
         Top             =   1680
         Width           =   225
      End
      Begin VB.Label lblColour 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   5
         Left            =   960
         TabIndex        =   13
         Top             =   1680
         Width           =   225
      End
      Begin VB.Label lblColour 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   12
         Top             =   2040
         Width           =   225
      End
      Begin VB.Label lblColour 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   7
         Left            =   600
         TabIndex        =   11
         Top             =   2040
         Width           =   225
      End
      Begin VB.Label lblColour 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   8
         Left            =   960
         TabIndex        =   10
         Top             =   2040
         Width           =   225
      End
      Begin VB.Label lblColour 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   9
         Top             =   2400
         Width           =   225
      End
      Begin VB.Label lblColour 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   10
         Left            =   600
         TabIndex        =   8
         Top             =   2400
         Width           =   225
      End
      Begin VB.Label lblColour 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   11
         Left            =   960
         TabIndex        =   7
         Top             =   2400
         Width           =   225
      End
      Begin VB.Label lblColour 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   12
         Left            =   240
         TabIndex        =   6
         Top             =   2760
         Width           =   225
      End
      Begin VB.Label lblColour 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   13
         Left            =   600
         TabIndex        =   5
         Top             =   2760
         Width           =   225
      End
      Begin VB.Label lblColour 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   14
         Left            =   960
         TabIndex        =   4
         Top             =   2760
         Width           =   225
      End
      Begin VB.Label lblColour 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   15
         Left            =   600
         TabIndex        =   3
         Top             =   3120
         Width           =   225
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu lb 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuUndo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuViewHelp 
         Caption         =   "&View Help"
         Shortcut        =   ^H
      End
   End
End
Attribute VB_Name = "frmIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private a As Long, b As Long, c As Long, e As Long

Private Lin As Long, Changed As Long

Const small = 9
Const medium = 19
Const large = 29
Const valAdd = 10

Private upPoint As Integer


Private Sub Form_Load()
  
    Picture1.BackColor = &HC0C0C0
    Picture3.BackColor = &HC0C0C0
    Picture3.Height = 4810
    Picture3.Width = 4810
    ImageList1.MaskColor = &HC0C0C0
    For e = 0 To 15
        lblColour(e).BackColor = QBColor(e)
    Next e
  
    Picture10 = Picture1.Image
    upPoint = small
    
    LoadGrid
    mnuUndo.Enabled = False
  
End Sub

Private Sub LoadGrid()

' Draws grid
    
    For F = 0 To Picture3.ScaleHeight Step 10
        Picture3.Line (0, F)-(Picture3.ScaleWidth, F), backCk
    Next F
    For F = 0 To Picture3.ScaleWidth Step 10
        Picture3.Line (F, 0)-(F, Picture3.ScaleHeight), backCk
    Next F
   
End Sub

Private Sub lblcolour_Click(Index As Integer)

'Choose Paint colour

    For a = 0 To 15
            lblColour(a).BorderStyle = 0
    Next a
  
    lblColour(Index).BorderStyle = 1
    a = QBColor(Index)
    Picture5.BackColor = a
  
End Sub

Private Sub mnuNew_Click()

'Create new icon

    savecheck
    a = &HC0C0C0
    Picture1.BackColor = &HC0C0C0
    Picture1.Width = 480
    Picture1.Height = 480
    Picture10 = Picture1.Image
    Fill
    LoadGrid
    Changed = 0
    mnuUndo.Enabled = False
  
End Sub

Private Sub mnuOpen_Click()

'open icon dialogue

    savecheck
    CommonDialog1.CancelError = True
    On Error GoTo error1
    CommonDialog1.FileName = ""
    CommonDialog1.Flags = cdlOFNFileMustExist
    CommonDialog1.Filter = "Icons (*.ico)|*.ico"
    CommonDialog1.ShowOpen
    
    If FileLen(CommonDialog1.FileName) <> 766 Then
        MsgBox "Ivalid or unsupported file format.", vbCritical
        Exit Sub
    End If

    MousePointer = 11
    Picture1.BackColor = &HC0C0C0
    Picture1 = LoadPicture(CommonDialog1.FileName)
    Picture3.PaintPicture Picture1.Image, 0, 0, 321, 321
    Changed = 1
    MousePointer = 0
    LoadGrid
    Exit Sub

error1:

End Sub

Private Sub mnuSave_Click()

'Save dialogue

    CommonDialog1.CancelError = True
    On Error GoTo error1
    CommonDialog1.FileName = ""
    CommonDialog1.Flags = cdlOFNOverwritePrompt + cdlOFNNoReadOnlyReturn
    CommonDialog1.Filter = "Icons (*.ico)|*.ico|Bitmaps (*.bmp)|*.bmp"
    CommonDialog1.ShowSave

    If CommonDialog1.FilterIndex = 1 Then

        MousePointer = 11
        Dim imgX As ListImage
        Set imgX = ImageList1.ListImages. _
        Add(1, , Picture1.Image)
        Dim picX As Picture
        Set picX = ImageList1.ListImages(1).ExtractIcon
        SavePicture picX, CommonDialog1.FileName
        Changed = 0
        MousePointer = 0
    
    End If

    If CommonDialog1.FilterIndex = 2 Then

        SavePicture Picture1.Image, CommonDialog1.FileName
        Changed = 0
    
    End If

    Exit Sub

error1:

End Sub

Private Function savecheck()

'Check if image has changed and prompt to save if it has

    On Error GoTo error1

    If Changed <> 0 Then
        Dim Msg, Style, Resp
        Msg = "Would you like to save your icon?"
        Style = vbYesNoCancel + vbExclamation
        Resp = MsgBox(Msg, Style)
    If Resp = vbYes Then mnuSave_Click: If Changed <> 0 Then Cancel = True
    If Resp = vbNo Then Cancel = False
    If Resp = vbCancel Then Cancel = True
    
    End If
    
error1:
End Function

Private Sub Form_Unload(Cancel As Integer)
    savecheck
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuUndo_Click()

'Undo drawing

    Picture1 = Picture6.Image
    Picture3.PaintPicture Picture4.Image, 0, 0, 321, 321
    LoadGrid
    mnuUndo.Enabled = False

End Sub

Private Sub mnuViewHelp_Click()
    frmHelp.Show
End Sub


Private Sub Option1_Click(Index As Integer)

'Selects the brush size

    If Index = 0 Then
        upPoint = small
    End If
    If Index = 1 Then
        upPoint = medium
    End If
        If Index = 2 Then
        upPoint = large
    End If
    
End Sub

Private Sub Picture3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Single pixel drawing

    Picture4 = Picture3.Image
    Picture6 = Picture1.Image
    b = Picture3.Point(X, Y)
    Changed = 1
    
    If Button = vbRightButton Then
     a = &HC0C0C0
    End If
    If Button = vbLeftButton Then
    a = Picture5.BackColor
    End If
    
    
    Lin = 1
    X1 = 0:  Y1 = 0
    For j = 0 To 31
        For P = 0 To 31
            
            If upPoint = 9 Then
                If X < X1 + valAdd And X > X1 And Y < Y1 + valAdd And Y > Y1 Then
                    Picture3.Line (X1 + 1, Y1 + 1)-(X1 + upPoint, Y1 + upPoint), a, BF
                    Picture1.PSet (X1 / valAdd, Y1 / valAdd), a
                End If
            End If
        
            If X < X1 + upPoint And X > X1 And Y < Y1 + upPoint And Y > Y1 Then
                Picture3.Line (X1 + 1, Y1 + 1)-(X1 + valAdd, Y1 + valAdd), a, BF
                Picture1.PSet (X1 / valAdd, Y1 / valAdd), a
            End If
        X1 = X1 + valAdd
        If X1 = 320 Then
        X1 = 0
        Y1 = Y1 + valAdd
    End If
        Next P
    Next j
  
    mnuUndo.Enabled = True
    LoadGrid
    
End Sub

Private Sub Picture3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Multiple pixel drawing, continuous

    'Displays mouse coordinates
    XPos = (X / 10) Mod 32 + 1
    YPos = (Y / 10) Mod 32 + 1
    Label1 = "Pos: " & CStr(XPos) & ", " & CStr(YPos)
    
    
    If Lin = 0 Then Exit Sub
    X1 = 0:  Y1 = 0
    For j = 0 To 31
  
        For P = 0 To 31
            
            If upPoint = 9 Then
                If X < X1 + valAdd And X > X1 And Y < Y1 + valAdd And Y > Y1 Then
                    Picture3.Line (X1 + 1, Y1 + 1)-(X1 + upPoint, Y1 + upPoint), a, BF
                    Picture1.PSet (X1 / valAdd, Y1 / valAdd), a
                End If
            End If
            If X < X1 + upPoint And X > X1 And Y < Y1 + upPoint And Y > Y1 Then
                Picture3.Line (X1 + 1, Y1 + 1)-(X1 + valAdd, Y1 + valAdd), a, BF
                Picture1.PSet (X1 / valAdd, Y1 / valAdd), a
                
            End If
            
    X1 = X1 + valAdd
    If X1 = 320 Then
    X1 = 0
    Y1 = Y1 + valAdd
    
    End If
    Next P
    Next j
    LoadGrid
    
End Sub

Private Sub Picture3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

' Stops drawing when mouse button is realeased

    Lin = 0
  
    If Button = vbLeftButton Then
        Picture10 = Picture1.Image
    End If
  
End Sub

Private Sub Fill()

'Fill image and preview with selected colour

    Picture4 = Picture3.Image
    Picture6 = Picture1.Image
  
    Changed = 1
  
    Picture1.Line (0, 0)-(31, 31), a, BF
    Picture3.BackColor = a
    LoadGrid
    Picture10 = Picture1.Image
    mnuUndo.Enabled = True
  
End Sub

Private Sub cmdFill_Click()
    Fill
End Sub


