VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.UserControl CDDB 
   ClientHeight    =   600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   870
   InvisibleAtRuntime=   -1  'True
   Picture         =   "CDDB.ctx":0000
   ScaleHeight     =   600
   ScaleWidth      =   870
   ToolboxBitmap   =   "CDDB.ctx":088B
   Windowless      =   -1  'True
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   60
      Top             =   60
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "CDDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

Public Event AllServerMessages(ByVal Text As String)
Public Event Connected()
Attribute Connected.VB_Description = "Fires when CDDB connects."
Public Event Disconnected()
Attribute Disconnected.VB_Description = "Fires when CDDB disconnects."
Public Event Error(ByVal Number As Long, ByVal Message As String)

'CDDB server sites
Public Enum CDDB_Server
    CDDB_None = 0   'do not connect
    CDDB_Random_US_Site = 1
    CDDB_San_Jose_CA_US = 2
    CDDB_Santa_Clara_CA_US = 3
End Enum

Private Const BASE_TRACK_NAME   As String = "track"
Private Const EXT_CDA           As String = ".cda"
Private Const TIME_FORMAT       As String = "hh:mm:ss"
Private Const SN_OFFSET         As Long = 24
Private Const SN_LEN            As Long = 4
Private Const TRACKLEN_OFFSET   As Long = 40
Private Const TRACKLEN_LEN      As Long = 4
Private Const TRACKNUM_OFFSET   As Long = 23
Private Const TRACKNUM_LEN      As Long = 1

Private fso                 As FileSystemObject
Private fsoDrive            As Drive
Private fsoFiles            As Files
Private udtCDDBServer       As CDDB_Server
Private strDriveLetter      As String
Private intTracks           As Integer
Private bytArray()          As Byte
Private strQueryString      As String
Private strGenre            As String
Private strTemp()           As String
Private strTemp2()          As String
Private strTempString       As String
Private strDiskID           As String
Private strArtistName       As String
Private strAlbumName        As String
Private lngAlbumLength      As Long
Private lngLeadoutOffset    As Long
Private colTrackNames       As Collection
Private colTrackLengths     As Collection
Private colTrackOffsets     As Collection
Private colTrackFrames      As Collection
Private blnDone             As Boolean

Private Sub Reset()
'Clears all data storage
    
    Dim dteTime As Date
    Dim i As Integer
    
    strArtistName = ""
    strAlbumName = ""
    strDiskID = ""
    strDriveLetter = ""
    strQueryString = ""
    strTempString = ""
    ReDim strTemp(0)
    ReDim strTemp2(0)
    intTracks = 0
    udtCDDBServer = CDDB_None
    
    'Make sure Winsock is closed - this resolved another Winsock issue
    dteTime = DateAdd("s", -1, Now) 'wait 1 sec for winsock to close
    If Winsock1.State <> sckClosed Then
        Winsock1.Close
        While Winsock1.State <> sckClosed And dteTime < Now
            DoEvents
        Wend
    End If
    
    'Clear collections
    ClearTrackNames
    For i = 1 To colTrackLengths.Count
        colTrackLengths.Remove 1
    Next i
    For i = 1 To colTrackOffsets.Count
        colTrackOffsets.Remove 1
    Next i
    For i = 1 To colTrackFrames.Count
        colTrackFrames.Remove 1
    Next i
    
End Sub

Private Sub ClearTrackNames()
    
    Dim i As Integer
    
    'Clear collection
    For i = 1 To colTrackNames.Count
        colTrackNames.Remove 1
    Next i

End Sub
Private Sub SetDefaultTrackNames()

    Dim i As Integer
    
    ClearTrackNames
    
    'Set default names
    For i = 1 To intTracks
        colTrackNames.Add BASE_TRACK_NAME & Format(i, "00") & EXT_CDA
    Next i
    
End Sub
Public Function SecondsToTimeString(Seconds As Long) As String
Attribute SecondsToTimeString.VB_Description = "Returns the a string in ""mm:ss"" format for the specified number of seconds."
'This is a utility function included to convert the total
'seconds to a time format

    Dim lngSec As Long
    Dim lngMin As Long
    
    lngMin = Int(Seconds / 60)
    lngSec = Seconds Mod 60
    
    SecondsToTimeString = lngMin & ":" & Format(lngSec, "00")
    
End Function

Public Property Get IsReady() As Boolean
Attribute IsReady.VB_Description = "Returns True if the CD is in the specified drive and it is an audio CD."
    
    If Not fsoDrive Is Nothing Then
        IsReady = fsoDrive.IsReady
    Else
        IsReady = False
    End If

End Property

Public Property Get CDDBServer() As CDDB_Server
    CDDBServer = udtCDDBServer
End Property

Public Property Get Genre() As String
Attribute Genre.VB_Description = "Returns the genre of the album."
    Genre = strGenre
End Property

Public Property Get AlbumName() As String
Attribute AlbumName.VB_Description = "Returns the name of the Album."
    AlbumName = strAlbumName
End Property

Public Property Get ArtistName() As String
Attribute ArtistName.VB_Description = "Returns the artist's name."
    ArtistName = strArtistName
End Property

Public Property Get CurrentDriveLetter() As String
Attribute CurrentDriveLetter.VB_Description = "Returns/Sets the current drive with the CD that is being processed."
    CurrentDriveLetter = strDriveLetter
End Property
    
Public Property Get TrackCount() As Long
Attribute TrackCount.VB_Description = "Returns the number of tracks on the CD."
    TrackCount = intTracks
End Property

Public Property Get DiskID() As String
Attribute DiskID.VB_Description = "Returns the serial number of the CD."
    DiskID = strDiskID
End Property

Public Property Get LeadoutOffset() As Long
Attribute LeadoutOffset.VB_Description = "Returns the frame position of the lead-out"
    LeadoutOffset = lngLeadoutOffset
End Property

Public Property Get TotalLength() As Long
Attribute TotalLength.VB_Description = "Returns the total length of the CD in seconds."
    TotalLength = lngAlbumLength
End Property

Public Property Get TrackName(Index As Integer) As String
Attribute TrackName.VB_Description = "Returns the name of the specified track."
    
    'Verify index
    If Index < 1 Or Index > colTrackNames.Count Then
        TrackName = ""
    Else
        TrackName = colTrackNames(Index)
    End If
    
End Property

Public Property Get TrackLength(Index As Integer) As Long
Attribute TrackLength.VB_Description = "Returns the time of the specified track in seconds."
    
    'Verify index
    If Index < 1 Or Index > colTrackLengths.Count Then
        TrackLength = 0
    Else
        TrackLength = Val(colTrackLengths(Index))
    End If
    
End Property

Public Property Get TrackOffset(Index As Integer) As Long
Attribute TrackOffset.VB_Description = "Returns the frame offset for the specified track."
    
    'Verify index
    If Index < 1 Or Index > colTrackOffsets.Count Then
        TrackOffset = 0
    Else
        TrackOffset = colTrackOffsets(Index)
    End If
    
End Property

Public Property Get TrackFrames(Index As Integer) As Long
Attribute TrackFrames.VB_Description = "Returns the length in frames of the specified track."
    
    'Verify index
    If Index < 1 Or Index > colTrackFrames.Count Then
        TrackFrames = 0
    Else
        TrackFrames = colTrackFrames(Index)
    End If
    
End Property

Private Sub UserControl_Resize()

    UserControl.Width = 870
    UserControl.Height = 600
    
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)

    Dim strText As String
    Dim a As Long
    
    'Initialize CDDB communication
    Winsock1.GetData strText, vbString
    RaiseEvent AllServerMessages(strText)
    If InStr(1, strText, "200 Hello and welcome ", vbTextCompare) Then
        Winsock1.SendData strQueryString & vbCrLf
    End If
    
    'Parse for music category
    If InStr(1, strText, "200 ", vbTextCompare) Then
        If InStr(1, strText, " / ", vbTextCompare) Then
            strTempString = InStr(4, strText, " ", vbTextCompare) + 1
            strGenre = Mid(strText, strTempString, InStr(strTempString, strText, " ", vbTextCompare) - strTempString)
            Winsock1.SendData "cddb read " & strGenre & " " & strDiskID & vbCrLf
        End If
    End If
    
    '???
    If InStr(1, strText, "211 Found inexact matches, list follows (until terminating `.')", vbTextCompare) Then
        strTemp = Split(strText, vbCrLf)
        strTempString = InStr(1, strTemp(1), " ", vbTextCompare) + 1
        strTempString = Mid(strTemp(1), strTempString, InStr(strTempString, strTemp(1), " ", vbTextCompare) - strTempString)
        strDiskID = strTempString
        strGenre = Left(strTemp(1), InStr(1, strTemp(1), " ", vbTextCompare) - 1)
        Winsock1.SendData "cddb read " & strGenre & " " & strDiskID & vbCrLf
        strTempString = ""
    End If
    
    'Couldn't find CD id in CDDB - disconnect
    If InStr(1, strText, CDDB_ERR_401, vbTextCompare) Then
        Winsock1.SendData "quit" & vbCrLf
        RaiseEvent Error(401, PRE_ERR_CDDB & CDDB_ERR_401)
    End If
    
    'Unrecognized command
    If InStr(1, strText, CDDB_ERR_500, vbTextCompare) Then
        Winsock1.SendData "quit" & vbCrLf
        RaiseEvent Error(500, PRE_ERR_CDDB & CDDB_ERR_500)
    End If
    
    'Invalid CD id - disconnect
    If InStr(1, strText, CDDB_ERR_501, vbTextCompare) Then
        Winsock1.SendData "quit" & vbCrLf
        RaiseEvent Error(501, PRE_ERR_CDDB & CDDB_ERR_501)
    End If
    
    'Parse for artist/album/track names
    If InStr(1, strText, " CD database entry follows (until terminating `.')", vbTextCompare) Then
        strTemp = Split(strText, vbCrLf)
        
        For a = 0 To UBound(strTemp)
            'Disk title
            If InStr(1, strTemp(a), "DTITLE=", vbTextCompare) Then
                strTempString = Mid(strTemp(a), 8, Len(strTemp(a)) - 7)
                strTemp2$ = Split(strTempString, " / ")
                strArtistName = strTemp2$(0)
                strAlbumName = strTemp2$(1)
                strTempString = ""
            End If
            
            'Track title
            If InStr(1, strTemp(a), "TTITLE", vbTextCompare) Then
                strTempString = InStr(1, strTemp(a), "=", vbTextCompare) + 1
                strTempString = Mid(strTemp(a), strTempString, Len(strTemp(a)) - strTempString + 1)
                Winsock1.SendData "quit" & vbCrLf
                colTrackNames.Add strTempString
                strTempString = ""
            End If
        Next
    End If

End Sub

Private Sub Winsock1_Close()
    
    blnDone = True
    RaiseEvent Disconnected

End Sub

Public Sub Cancel()
    
    Winsock1.SendData "quit" & vbCrLf
    blnDone = True

End Sub

Public Function ProcessCD(ByVal DriveLetter As String, Optional ByVal Server As CDDB_Server = CDDB_None) As Boolean
Attribute ProcessCD.VB_Description = "Retrieves CDDB information for the specified drive (CD) and CDDB server.  Returns True if successful."

    Dim strServerAddr As String
    Dim i As Integer
    
    On Error GoTo Error_Handler
    
    Reset
    
    udtCDDBServer = Server
    strDriveLetter = DriveLetter
    
    'Check drive letter
    If Trim(DriveLetter) = "" Then
        Err.Raise 1005, , "Please specify a drive letter"
    End If
    
    'Make sure the drive exists
    If fso.DriveExists(DriveLetter) Then
        Set fsoDrive = fso.GetDrive(DriveLetter)
        strDriveLetter = fsoDrive.Path
        DriveLetter = strDriveLetter
    Else
        Err.Raise 1000, "Drive " & DriveLetter & " does not exist"
    End If
    
    'Make sure the drive is a CD drive
    If fso.GetDrive(DriveLetter).DriveType <> CDRom Then
        Err.Raise 1001 + vbError, , "Drive " & DriveLetter & " is not a CD-Rom drive"
    End If
    
    'Make sure there is a CD in the drive
    If fso.GetDrive(DriveLetter).IsReady Then
        Set fsoFiles = fsoDrive.RootFolder.Files
        intTracks = fsoFiles.Count
    Else
        Err.Raise 1002 + vbError, , "There is no CD in drive " & DriveLetter
    End If
    
    'Make sure we have an audio CD in the drive
    If Not fso.FileExists(fso.BuildPath(DriveLetter, BASE_TRACK_NAME & "01" & EXT_CDA)) Then
        Err.Raise 1003 + vbError, , "The CD in drive " & DriveLetter & " is not an audio CD"
    End If
    
    'Set some properties
    Dim o As New CCd
    If Not o.Init(DriveLetter) Then
        Err.Raise 1010, , PRE_ERR_DEVICE & Err.Description
    End If
    SetDefaultTrackNames
    strDiskID = o.DiskID
    lngAlbumLength = o.TotalLength
    lngLeadoutOffset = o.TrackOffset(o.TrackCount)
    
    'Class CCd Track times cumulative - "unaccumulate" them
    For i = 1 To o.TrackCount
        colTrackOffsets.Add o.TrackOffset(i - 1)
        colTrackLengths.Add o.TrackTime(i) - o.TrackTime(i - 1)
        colTrackFrames.Add o.TrackFrames(i - 1)
    Next i
    
    '"Fix" last track time
    colTrackLengths.Remove o.TrackCount
    colTrackLengths.Add o.TotalLength - o.TrackTime(o.TrackCount - 1)
    
    'Exit if we don't want CDDB info
    If Server = CDDB_None Then
        ProcessCD = True
        Exit Function
    End If
    
    'Make way for real track names
    ClearTrackNames
    
    'Set the server string
    Select Case Server
        Case CDDB_Random_US_Site
            strServerAddr = CDDB_SERVER_RANDOM_US
        Case CDDB_San_Jose_CA_US
            strServerAddr = CDDB_SERVER_SJ_CA
        Case CDDB_Santa_Clara_CA_US
            strServerAddr = CDDB_SERVER_SC_CA
        Case Else
            strServerAddr = ""
    End Select
    
    'Contact CDDB & set the rest of the properties
    strQueryString = o.QueryString
    Winsock1.LocalPort = CDDB_PORT
    Winsock1.RemotePort = CDDB_PORT
    Winsock1.RemoteHost = strServerAddr
    Winsock1.Connect
    Set o = Nothing
    
    'Wait until winsock is done
    blnDone = False
    While Not blnDone 'gets set in Winsock1_Close or sub Cancel
        DoEvents
    Wend
    
    ProcessCD = True

Exit Function

Error_Handler:
    RaiseEvent Error(Err.Number, Err.Description)
    ProcessCD = False
    
End Function

Private Sub Winsock1_Connect()

    RaiseEvent Connected
    Winsock1.SendData "cddb hello " & Winsock1.LocalHostName & " " & Winsock1.LocalIP & App.Title & vbCrLf
    
End Sub

Public Sub About()
Attribute About.VB_Description = "Shows the About information screen."
Attribute About.VB_UserMemId = -552
Attribute About.VB_MemberFlags = "40"
    frmAbout.Show vbModal
End Sub

Private Sub UserControl_Initialize()
    
    Set fso = New FileSystemObject
    Set colTrackNames = New Collection
    Set colTrackLengths = New Collection
    Set colTrackFrames = New Collection
    Set colTrackOffsets = New Collection

End Sub

Private Sub UserControl_Terminate()
    
    Set fso = Nothing
    Set fsoDrive = Nothing
    Set colTrackNames = Nothing
    Set colTrackLengths = Nothing

End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

    Winsock1.Close
    
    If Number = 10048 Then
        'Seemingly solid workaround for MSDN problem
        'PRB: Winsock Control Generates Error 10048 - Address in Use
        ProcessCD strDriveLetter, udtCDDBServer
    Else
        blnDone = True
        RaiseEvent Error(Number, PRE_ERR_WINSOCK & Description)
    End If
    
End Sub
