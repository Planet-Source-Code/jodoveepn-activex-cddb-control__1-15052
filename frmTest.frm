VERSION 5.00
Object = "*\ACDDBControl.vbp"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   8070
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6000
      TabIndex        =   20
      Top             =   3120
      Width           =   855
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   5760
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   2520
      Width           =   2175
   End
   Begin VB.TextBox txtDrive 
      Height          =   285
      Left            =   4920
      MaxLength       =   1
      TabIndex        =   16
      Top             =   2520
      Width           =   735
   End
   Begin VB.TextBox txtSerialNumber 
      Height          =   285
      Left            =   4920
      TabIndex        =   13
      Top             =   360
      Width           =   2175
   End
   Begin MSComctlLib.ListView lvwTracks 
      Height          =   1695
      Left            =   120
      TabIndex        =   12
      Top             =   3840
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   2990
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "#"
         Object.Width           =   176
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Offset"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Frames"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Time"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Name"
         Object.Width           =   882
      EndProperty
   End
   Begin VB.TextBox txtGenre 
      Height          =   285
      Left            =   7200
      TabIndex        =   11
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox txtTime 
      Height          =   285
      Left            =   7200
      TabIndex        =   9
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox txtArtist 
      Height          =   285
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1080
      Width           =   2175
   End
   Begin VB.TextBox txtAlbum 
      Height          =   285
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1800
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Query..."
      Default         =   -1  'True
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   7080
      TabIndex        =   1
      Top             =   3120
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   3135
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   360
      Width           =   4575
   End
   Begin VB.Label Label9 
      Caption         =   "CDDB Server"
      Height          =   255
      Left            =   5760
      TabIndex        =   18
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "CDDB Messages"
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label7 
      Caption         =   "Drive"
      Height          =   255
      Left            =   4920
      TabIndex        =   15
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label Label6 
      Caption         =   "Disk ID"
      Height          =   255
      Left            =   4920
      TabIndex        =   14
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "Genre"
      Height          =   255
      Left            =   7200
      TabIndex        =   10
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "Time"
      Height          =   255
      Left            =   7200
      TabIndex        =   8
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "Tracks"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Artist"
      Height          =   255
      Left            =   4920
      TabIndex        =   5
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Album"
      Height          =   255
      Left            =   4920
      TabIndex        =   3
      Top             =   1560
      Width           =   615
   End
   Begin CDDBControl.CDDB CDDB1 
      Left            =   7200
      Top             =   120
      _ExtentX        =   1535
      _ExtentY        =   1058
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private blnFinished As Boolean

Private Sub CDDB1_AllServerMessages(ByVal Text As String)
    Text1.Text = Text1.Text & Text
End Sub

Private Sub CDDB1_Connected()
    Text1.Text = "*** Connected ***" & vbCrLf & vbCrLf
End Sub

Private Sub CDDB1_Disconnected()
    blnFinished = True
    Text1.Text = Text1 & vbCrLf & "*** Disconnected ***" & vbCrLf
End Sub

Private Sub CDDB1_Error(ByVal Number As Long, ByVal Message As String)

    MsgBox Message
    blnFinished = True
    
End Sub

Private Sub cmdExit_Click()

    blnFinished = True 'just in case
    Unload Me
    
End Sub

Private Sub Command1_Click()
    CDDB1.Cancel
End Sub

Private Sub Command2_Click()
    
    Dim strTemp() As String
    Dim strDrive As String
    Dim i As Integer
    Dim lstItem As ListItem
    
    Command2.Enabled = False
    Command1.Enabled = True
        
    Text1.Text = ""
    
    'Process the CD
    blnFinished = False
    strDrive = Left(txtDrive.Text, 1)
    If Not CDDB1.ProcessCD(strDrive, Combo1.ListIndex) Then
        MsgBox "Process failed"
    End If
    
    Command2.Enabled = True
    
    'Build arrays for ListView control
    ReDim strTemp(CDDB1.TrackCount, 4)
    For i = 1 To CDDB1.TrackCount
        strTemp(i, 0) = i
        strTemp(i, 1) = CDDB1.TrackOffset(i)
        strTemp(i, 2) = CDDB1.TrackFrames(i)
        strTemp(i, 3) = CDDB1.SecondsToTimeString(CDDB1.TrackLength(i))
        strTemp(i, 4) = CDDB1.TrackName(i)
    Next i
    
    'Display album info
    lvwTracks.ListItems.Clear
    txtAlbum.Text = CDDB1.DiskID
    txtAlbum.Text = CDDB1.AlbumName
    txtArtist.Text = CDDB1.ArtistName
    txtTime.Text = CDDB1.SecondsToTimeString(CDDB1.TotalLength)
    txtGenre.Text = CDDB1.Genre
    txtSerialNumber.Text = CDDB1.DiskID
    For i = 1 To CDDB1.TrackCount
        Set lstItem = lvwTracks.ListItems.Add(, , strTemp(i, 0))
        lstItem.SubItems(1) = strTemp(i, 1)
        lstItem.SubItems(2) = strTemp(i, 2)
        lstItem.SubItems(3) = strTemp(i, 3)
        lstItem.SubItems(4) = strTemp(i, 4)
    Next i
    
    'Clean up
    Command1.Enabled = False
    Erase strTemp
    Set lstItem = Nothing
    
End Sub

Private Sub Form_Load()
    
    Me.Show
    
    'Setup listview column widths
    lvwTracks.ColumnHeaders(1).Width = 300
    lvwTracks.ColumnHeaders(2).Width = 1200
    lvwTracks.ColumnHeaders(3).Width = 1200
    lvwTracks.ColumnHeaders(4).Width = 600
    lvwTracks.ColumnHeaders(5).Width = 4100
    
    'Load combo box
    Combo1.AddItem "None"
    Combo1.AddItem "Random US"
    Combo1.AddItem "San Jose, CA"
    Combo1.AddItem "Santa Clara, CA"
    Combo1.ListIndex = 0
    
    Command1.Enabled = False
    
    txtDrive.SetFocus
    
End Sub
