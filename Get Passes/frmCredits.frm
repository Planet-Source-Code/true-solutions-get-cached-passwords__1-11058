VERSION 5.00
Begin VB.Form frmCredits 
   BackColor       =   &H00C00000&
   BorderStyle     =   0  'None
   Caption         =   "Top to Bottom Credits"
   ClientHeight    =   9000
   ClientLeft      =   4530
   ClientTop       =   4875
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      Height          =   9015
      Left            =   0
      ScaleHeight     =   9015
      ScaleWidth      =   11895
      TabIndex        =   0
      Top             =   0
      Width           =   11895
      Begin VB.Timer Timer1 
         Interval        =   10
         Left            =   4080
         Top             =   600
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H0000FF00&
         Height          =   9015
         Left            =   0
         Picture         =   "frmCredits.frx":0000
         ScaleHeight     =   9015
         ScaleWidth      =   11895
         TabIndex        =   1
         Top             =   0
         Width           =   11895
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Symbol"
               Size            =   8.25
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   165
            Left            =   1080
            TabIndex        =   4
            Top             =   3480
            Width           =   1815
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "This Program was coded by Scottybee"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   855
            Left            =   960
            TabIndex        =   2
            Top             =   0
            Width           =   3495
         End
      End
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "z"
      BeginProperty Font 
         Name            =   "Symbol"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   150
   End
End
Attribute VB_Name = "frmCredits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Click()
End
End Sub

Private Sub Label5_Click()
End
End Sub

Private Sub Picture1_Click()
End
End Sub

Private Sub Timer1_Timer()
If Picture2.Top < Picture1.Height - Picture1.Height - Picture2.Height Then
    Picture2.Top = Picture2.Height - 1
    
    Picture2.Top = Label2.Top - 10
    
Else
    Picture2.Top = Picture2.Top - 10
    
End If

End Sub
