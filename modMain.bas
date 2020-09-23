Attribute VB_Name = "modMain"
Option Explicit
'======================================================
'
'CDDB ActiveX Control For VB5/VB6
'Compiled With VB6 SP3
'Note: You will NEED VB6 to compile this again.
'      Because I used a VB6 Function 'Split'
'      For more info, open your MSDN Help (F1)
'      and jump to this URL:
'
'
'mk:@MSITStore:C:\Program%20Files\Microsoft%20Visual%20Studio\MSDN\2000APR\1033\vbenlr98.chm::/html/vafctSplit.htm
'      Or, paste that in your browser. If the path
'      Doesn't match your path. It doesn't matter
'      It will still find it.
'June 24, 2000
'Known Bugs? Not sure, I won't be using this control
'I only remade it because it was a request.
'There might be errors, I redid this whole control
'from the start. Took less then a day to make.
'Author: Michael L. Barker
'
'======================================================
'Version 1.1
'Feb. 2, 2001 - Upgraded by Todd P. Worland - jodoveepn@yahoo.com
'
'Made the following changes:
'
'   -   Waits for CDDB to finish communication before releasing control
'   -   Added sub Cancel to cancel CCDB communication and release control
'           in case of slow or problem connection
'   -   Added a fixes to allow CDDB to connect multiple times within a short
'           time span
'   -   Broke up Author/Artist event into properties
'   -   Exposed most properties in the Class (CCd.cls) that were
'       hidden from the control
'   -   Added Error event & error handling
'   -   Added toolbar icon
'   -   Other minor code changes
'   -   Cleaned up code & added more comments
'   -   Added a better (in my opinion) sample App
'   -   Added a nifty About screen (flickery but quick and easy)
'
'======================================================

Public Const CDDB_PORT              As Long = 8880

Public Const CDDB_SERVER_RANDOM_US  As String = "us.cddb.com"
Public Const CDDB_SERVER_SJ_CA      As String = "sj.ca.us.cddb.com"
Public Const CDDB_SERVER_SC_CA      As String = "sc.ca.us.cddb.com"

'Some CDDB Errors
Public Const CDDB_ERR_401           As String = "401 No such CD entry in database"
Public Const CDDB_ERR_500           As String = "500 Unrecognized command"
Public Const CDDB_ERR_501           As String = "501 Invalid disc ID"

'Message prefixes
Public Const PRE_ERR_CDDB           As String = "CDDB Error: "
Public Const PRE_ERR_DEVICE         As String = "CD-Rom Device Error: "
Public Const PRE_ERR_WINSOCK        As String = "Winsock Error: "

