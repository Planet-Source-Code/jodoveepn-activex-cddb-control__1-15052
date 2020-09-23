VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4590
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   4590
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2640
      Top             =   3240
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   3240
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      Height          =   3015
      Left            =   120
      ScaleHeight     =   2955
      ScaleWidth      =   4275
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin VB.PictureBox Picture2 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   600
         Left            =   120
         Picture         =   "frmAbout.frx":08CA
         ScaleHeight     =   600
         ScaleWidth      =   870
         TabIndex        =   3
         Top             =   120
         Width           =   870
      End
      Begin VB.Label lblAbout 
         Alignment       =   2  'Center
         Caption         =   "lblText"
         Height          =   3735
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   4095
      End
   End
   Begin VB.Label Label1 
      Caption         =   "This control is Freeware - enjoy"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3240
      Width           =   2415
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DISPLACEMENT As Integer = 20
Private strText As String

Private Sub cmdOk_Click()
    Unload Me
End Sub

Private Sub Form_Load()

    Me.Caption = "About " & App.Title
    
    'Start scrolling from bottom of the picture box
    lblAbout.Top = Picture1.Top + Picture1.Height
    
    'Setup the label
    strText = "= = = = = = = = = = = = = = = = = =" & vbCrLf & _
        App.Title & vbCrLf & "Version " & _
        App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & _
        "= = = = = = = = = = = = = = = = = =" & vbCrLf & vbCrLf & _
        "Thanks to the following people" & vbCrLf & _
        "who wrote most of this code:" & vbCrLf & vbCrLf & _
        "----- Michael L. Barker -----" & vbCrLf & _
        "for supplying the original" & vbCrLf & _
        "CDDB control (version 1.0.0)" & vbCrLf & vbCrLf & _
        "----- Anonymous -----" & vbCrLf & _
        "for the CDDB protocol class (CCd.cls)" & vbCrLf & vbCrLf & _
        "= = = = = = = = = = = = = = = = = =" & vbCrLf & vbCrLf & _
        "Upgraded by Todd P. Worland" & vbCrLf & _
        "jodoveepn@yahoo.com" & vbCrLf & vbCrLf & vbCrLf & _
        "See the readme.txt for specific version changes"
        
    lblAbout.Caption = strText
    
    'Start scrolling
    Timer1.Enabled = True
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Stop scrolling
    Timer1.Enabled = False
End Sub

Private Sub Timer1_Timer()

    'Move the label
    lblAbout.Top = lblAbout.Top - DISPLACEMENT
    If lblAbout.Top < -lblAbout.Height Then
        lblAbout.Top = Picture1.Top + Picture1.Height
    End If
    
End Sub
