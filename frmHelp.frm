VERSION 5.00
Begin VB.Form frmHelp 
   Caption         =   "Help"
   ClientHeight    =   6045
   ClientLeft      =   4590
   ClientTop       =   2595
   ClientWidth     =   6150
   LinkTopic       =   "Form2"
   ScaleHeight     =   6045
   ScaleWidth      =   6150
   Begin VB.CommandButton Command1 
      Caption         =   "Close Help"
      Height          =   495
      Left            =   2400
      TabIndex        =   3
      Top             =   5280
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   2400
      Picture         =   "frmHelp.frx":0000
      ScaleHeight     =   1215
      ScaleWidth      =   1215
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Use the right mouse button to erase part of your icon without selecting a colour"
      Height          =   615
      Left            =   3960
      TabIndex        =   2
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Use the left mouse button to select a colour from the palette. Then simply draw your icon on the main canvas."
      Height          =   975
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   3420
      Left            =   1200
      Picture         =   "frmHelp.frx":31A6
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   3540
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me


End Sub
