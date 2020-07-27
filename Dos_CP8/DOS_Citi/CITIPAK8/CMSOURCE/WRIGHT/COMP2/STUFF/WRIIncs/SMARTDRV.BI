DECLARE SUB SDFlush ()
'Name: SDFLUSH
'Desc: Flush SmartDrive disk caches

DECLARE SUB SDVersion (MajorV%, MinorV%)
'Name: SDVersion
'Desc: Get SmartDrive version number

DECLARE FUNCTION SDRCached% (Drive$)
'Name: SDRCACHED
'Desc: Determine whether a drive is read cached by SmartDrive

DECLARE SUB SDReset ()
'Name: SDRESET
'Desc: Reset SmartDrive disk caches

DECLARE SUB SDStats (Hits&, Misses&)
'Name: SDSTATS
'Desc: Get SmartDrive cache hits and misses

DECLARE FUNCTION SDWCached% (Drive$)
'Name: SDWCACHED
'Desc: Determine whether a drive is write cached by SmartDrive

