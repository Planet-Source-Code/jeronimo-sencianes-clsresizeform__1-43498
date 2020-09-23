VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5940
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8550
   LinkTopic       =   "Form1"
   ScaleHeight     =   5940
   ScaleWidth      =   8550
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture7 
      BackColor       =   &H00FF80FF&
      Height          =   870
      Left            =   1290
      ScaleHeight     =   810
      ScaleWidth      =   5925
      TabIndex        =   7
      Top             =   5010
      Width           =   5985
   End
   Begin VB.PictureBox Picture6 
      BackColor       =   &H000080FF&
      Height          =   930
      Left            =   3795
      ScaleHeight     =   870
      ScaleWidth      =   915
      TabIndex        =   6
      Top             =   135
      Width           =   975
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00FFC0C0&
      Height          =   930
      Left            =   60
      ScaleHeight     =   870
      ScaleWidth      =   1155
      TabIndex        =   5
      Top             =   60
      Width           =   1215
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00FFC0C0&
      Height          =   1950
      Left            =   60
      ScaleHeight     =   1890
      ScaleWidth      =   1125
      TabIndex        =   4
      Top             =   3945
      Width           =   1185
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00FFC0C0&
      Height          =   1890
      Left            =   7290
      ScaleHeight     =   1830
      ScaleWidth      =   1155
      TabIndex        =   3
      Top             =   3990
      Width           =   1215
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFC0C0&
      Height          =   945
      Left            =   7320
      ScaleHeight     =   885
      ScaleWidth      =   1125
      TabIndex        =   2
      Top             =   60
      Width           =   1185
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FF80FF&
      Height          =   1425
      Left            =   1290
      ScaleHeight     =   1365
      ScaleWidth      =   5925
      TabIndex        =   1
      Top             =   3525
      Width           =   5985
   End
   Begin VB.PictureBox pic1 
      BackColor       =   &H00FF8080&
      Height          =   930
      Left            =   60
      ScaleHeight     =   870
      ScaleWidth      =   8385
      TabIndex        =   0
      Top             =   2565
      Width           =   8445
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oRZ As clsResizeForm

Private Sub Form_Load()
    Set oRZ = New clsResizeForm
    oRZ.FormName = Me
    oRZ.MinResizeX = 8600
    oRZ.MinResizeY = 6100
    
    oRZ.Add Picture1, 1
    oRZ.Add Picture2, 2
    oRZ.Add Picture3, 3
    oRZ.Add pic1, 4
    oRZ.Add Picture7, 5
    oRZ.Add Picture4, 6
    
    oRZ.Add Picture6, 9
End Sub

Private Sub Form_Resize()
    oRZ.ResizeAll
End Sub
