CDDB ActiveX Control
original by Michael L. Barker
upgraded by Todd P. Worland

compiled with VB6 sp4 and with VB7 (.NET) installed


===========
Description
===========
This control will expose CD information such as the number of tracks, track lengths (in seconds)/frames/offsets, and serial number.  It will also connect to CDDB and retrieve artist, album, and track names.

I took Michael's control and reworked it a little to suit my needs for a zestier CDDB control.  This control will now allow access to the following info bits:
	- CD Artist
	- Album name
	- Album genre
	- Album length (in seconds)
	- Track count
	- Track names
	- Track times (in seconds)
	- Track offsets
	- Track frames
	- Disk serial number (disk ID)
	- Lead out offset

Note that lengths are in seconds.  Use the method "SecondsToTimeString" to convert the value to a string in "mm:ss" format.

Lastly, becuase Winsock runs on its own thread, it's a bit tricky to code the communication behavior especially in the case of an error.  This is the main reason I coded an Error event.  Take note of the sample app and how it disconnects from the Winsock.


=======
Changes
=======
There are several fixes I put in to make it a fairy robust control:
	- It now waits for the end of CDDB communication (or sub Cancel) before
	  returning control to the user
	- Added a fix that makes sure Winsock is closed before connecting again
	- Added a workaround for the MSDN article:
	  "PRB: Winsock Control Generates Error 10048 - Address in Use"
	  Without this fix, this error would be generated upon attempting to
	  contact the same CDDB server a second time in a row

	  This error makes for some interesting reading on how ports are handled.
	  If you're really interested, do a search for SO_REUSEADDR in the MSDN library.


============
Requirements
============
There are 2 main pieces that you will need to get this control to work, both of which should be taken care of automatically, but just in case...
	- the winsock control 
	- the file system object reference (scrrun.dll)


============
Known Issues
============
	- During testing, sometimes exiting the program will unload the form but the code
	  will still be running.  I leave testing the compiled version to you.  :o/
	- The offsets returned from CDDB seem slightly different than the
	  ones that are calculated by this control
	- The frames that are calculated by this control seem suspiciously
	  inconsistent.  I'm not sure they're correct, and was too lazy to
	  verify it.  If you discover that the frame calculation is in fact
	  incorrect, I'd appreciate an email.
	- The leadout always seems to be 0.
	- No others (yet)...

========
Comments
========
This control is freeware.  Enjoy.

If you rework this control, have comments, or questions, please email me at jodoveepn@yahoo.com


=======
Credits
=======
Thanks to the following:
	- Michael L. Barker for the original code and control
	- Whoever posted the CDDB protocol code (CCd.cls - see Michael's comment)


====
TODO
====
	- Add an enum of all the types of genres
	- Add properties to expose track artists/names for CDs that have various artists
	  (process a soundtrack CD and see what I mean)
	- Investigate the frame/offset issue more.  I found an MSDN article
	  entitled "Getting Disc Information."  It contains code that may work
	  correctly?


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
'   -	Waits for CDDB to finish communication before releasing control
'   -	Added sub Cancel to cancel CCDB communication and release control
'	in case of slow or problem connection
'   -	Added a fixes to allow CDDB to connect multiple times within a short
'	time span
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
