VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Style Menu"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7020
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   7020
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picCarry 
      Height          =   3135
      Left            =   0
      ScaleHeight     =   3075
      ScaleWidth      =   915
      TabIndex        =   1
      Top             =   120
      Width           =   975
      Begin VB.CommandButton cmdMenu 
         Caption         =   "Menu"
         Height          =   315
         Left            =   0
         TabIndex        =   0
         TabStop         =   0   'False
         ToolTipText     =   "Menu"
         Top             =   2760
         Width           =   920
      End
      Begin VB.Image imgs 
         Height          =   600
         Left            =   1440
         Picture         =   "Form2.frx":0000
         Stretch         =   -1  'True
         Top             =   1080
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Image img 
         Height          =   360
         Index           =   0
         Left            =   240
         Picture         =   "Form2.frx":08CA
         Stretch         =   -1  'True
         ToolTipText     =   "ID Person"
         Top             =   240
         Width           =   360
      End
      Begin VB.Image img 
         Height          =   360
         Index           =   1
         Left            =   240
         Picture         =   "Form2.frx":1194
         Stretch         =   -1  'True
         ToolTipText     =   "Send fax"
         Top             =   840
         Width           =   360
      End
      Begin VB.Image img 
         Height          =   360
         Index           =   2
         Left            =   240
         Picture         =   "Form2.frx":1A5E
         Stretch         =   -1  'True
         ToolTipText     =   "Accessories"
         Top             =   1440
         Width           =   360
      End
      Begin VB.Image img 
         Height          =   360
         Index           =   3
         Left            =   240
         Picture         =   "Form2.frx":2328
         Stretch         =   -1  'True
         ToolTipText     =   "Personal"
         Top             =   2040
         Width           =   360
      End
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   3930
      TabIndex        =   2
      Top             =   480
      Width           =   75
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Dim exist As Boolean
Dim num As Integer
Const topBut = 2760
Const heightPic = 3135
Private Sub cmdMenu_Click()
Dim i As Integer
    If cmdMenu.Top <> 0 Then
        For i = cmdMenu.Top To 0 Step -1
            cmdMenu.Top = i
            picCarry.Height = picCarry.Height - 1
            DoEvents
        Next i
    Else
        For i = cmdMenu.Top To topBut
            cmdMenu.Top = i
            picCarry.Height = picCarry.Height + 1
            DoEvents
        Next i
    End If
End Sub

Private Sub Form_Load()
    lblMessage.Caption = "You like this?" & vbLf & "It simple, no need expert to do this" & vbLf & "just vote me"
    exist = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgs.Visible = False
End Sub

Private Sub img_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim tempx As Integer, tempy As Integer
    imgs.Visible = False
    num = Index
    tempx = img(Index).Left + Int(img(Index).Width / 2)
    tempy = img(Index).Top + Int(img(Index).Height / 2)
    imgs.Left = tempx - Int(imgs.Width / 2)
    imgs.Top = tempy - Int(imgs.Height / 2)
    imgs.Picture = img(Index).Picture
    sndPlaySound App.Path & "\boop.wav", 0
    imgs.Visible = True
    imgs.ToolTipText = img(Index).ToolTipText
End Sub

Private Sub imgs_Click()
    MsgBox "You have click the button #" & num + 1
End Sub

Private Sub picCarry_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgs.Visible = False
End Sub
