VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Split Name"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   3600
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtFullname2 
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Text            =   "Mr, Ankur, Arora, Sr"
      Top             =   2130
      Width           =   3255
   End
   Begin VB.CommandButton cmdSplit2 
      Caption         =   "Split Name"
      Default         =   -1  'True
      Height          =   345
      Left            =   1020
      TabIndex        =   5
      Top             =   2550
      Width           =   1575
   End
   Begin VB.CommandButton cmdSplit 
      Caption         =   "Split Name"
      Height          =   345
      Left            =   1020
      TabIndex        =   1
      Top             =   870
      Width           =   1575
   End
   Begin VB.TextBox txtFullname 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "Mr Ankur Arora Sr"
      Top             =   450
      Width           =   3255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Split Name According to Commas"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   1770
      Width           =   3255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "OR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   1380
      Width           =   3255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Split Name According to Spaces"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   90
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdSplit_Click()
    Module1.SplitName (txtFullname.Text)
    txtFullname.SetFocus
End Sub

Private Sub cmdSplit2_Click()
    Module1.SplitName (txtFullname2.Text)
    txtFullname2.SetFocus
End Sub

Private Sub txtFullname_GotFocus()
    txtFullname.SelStart = 0
    txtFullname.SelLength = (Len(txtFullname.Text))
End Sub


Private Sub txtFullname2_GotFocus()
    txtFullname2.SelStart = 0
    txtFullname2.SelLength = (Len(txtFullname2.Text))
End Sub
