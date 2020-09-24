VERSION 5.00
Begin VB.Form GetCachedPasses 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Get cached Passwords By Scottybee"
   ClientHeight    =   5445
   ClientLeft      =   720
   ClientTop       =   1455
   ClientWidth     =   10365
   ForeColor       =   &H00808080&
   Icon            =   "getcachedpasses.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "getcachedpasses.frx":030A
   ScaleHeight     =   5445
   ScaleWidth      =   10365
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.ListBox List1 
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2580
      ItemData        =   "getcachedpasses.frx":CD43
      Left            =   1440
      List            =   "getcachedpasses.frx":CD45
      TabIndex        =   0
      Top             =   960
      Width           =   8775
   End
   Begin VB.Image ImgSavePasses 
      Height          =   525
      Left            =   4680
      Picture         =   "getcachedpasses.frx":CD47
      Top             =   4560
      Width           =   2625
   End
End
Attribute VB_Name = "GetCachedPasses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
Call GetPasswords
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub ImgSavePasses_Click()
Call Save_ListBox("c:\windows\desktop\YourPasses.txt", List1)
End Sub

