VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Tetris VB6 2025"
   ClientHeight    =   10230
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   ScaleHeight     =   10230
   ScaleWidth      =   5880
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton box 
      BackColor       =   &H000000FF&
      Height          =   495
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Height          =   9615
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const box_size As Integer = 400

Private Sub Form_Load()
    Frame1.Width = 10 * box_size
    Frame1.Height = 24 * box_size
End Sub
