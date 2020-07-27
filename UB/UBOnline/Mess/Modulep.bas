Attribute VB_Name = "Module1"
Public NAML, NAMF As String
Public goingelsewhere As Boolean
Public DOPOPUP, nwl, nwc, nww, nwr, nws, nwb, nwi, nwj, nwm, usesf As String
Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Declare Function WriteProfileString Lib "kernel32" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lparam As String) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Integer, ByVal Y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal uFlags As Long) As Boolean
Public Const HWND_BROADCAST = &HFFFF
Public Const WM_WININICHANGE = &H1A
Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As OSVERSIONINFO) As Integer
Public Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, pDefault As PRINTER_DEFAULTS) As Long
Public Declare Function SetPrinter Lib "winspool.drv" Alias "SetPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, pPrinter As Any, ByVal Command As Long) As Long
Public Declare Function GetPrinter Lib "winspool.drv" Alias "GetPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, pPrinter As Any, ByVal cbBuf As Long, pcbNeeded As Long) As Long
Public Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Any) As Long
Public Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long      ' constants for DEVMODE structure
Public Const CCHDEVICENAME = 32
Public Const CCHFORMNAME = 32
' constants for DesiredAccess member of PRINTER_DEFAULTS
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const PRINTER_ACCESS_ADMINISTER = &H4
Public Const PRINTER_ACCESS_USE = &H8
Public Const PRINTER_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or _
PRINTER_ACCESS_ADMINISTER Or PRINTER_ACCESS_USE)

Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

' constant that goes into PRINTER_INFO_5 Attributes member
' to set it as default
Public Const PRINTER_ATTRIBUTE_DEFAULT = 4
Public Type DEVMODE
    dmDeviceName As String * CCHDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCHFORMNAME
    dmLogPixels As Integer
    dmBitsPerPel As Long
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
    dmICMMethod As Long        ' // Windows 95 only
    dmICMIntent As Long        ' // Windows 95 only
    dmMediaType As Long        ' // Windows 95 only
    dmDitherType As Long       ' // Windows 95 only
    dmReserved1 As Long        ' // Windows 95 only
    dmReserved2 As Long        ' // Windows 95 only
End Type
Public Type PRINTER_INFO_5
        pPrinterName As String
        pPortName As String
        Attributes As Long
        DeviceNotSelectedTimeout As Long
        TransmissionRetryTimeout As Long
End Type
Public Type PRINTER_DEFAULTS
        pDatatype As Long
        pDevMode As Long
        DesiredAccess As Long
End Type


Public Function RemoveSpecChars(strVal As String) As String
    For X% = 1 To Len(strVal)
        If InStr("!@#$%^&*><?`~", Mid$(strVal, X%, 1)) = 0 Then
            RemoveSpecChars = RemoveSpecChars & Mid$(strVal, X%, 1)
        End If
    Next X%
End Function
Public Function SetAlwaysOnTop(objForm As Form) As Boolean
objForm.Show
SetAlwaysOnTop = SetWindowPos(objForm.hwnd, HWND_TOPMOST, objForm.Left, objForm.Top, objForm.Width, objForm.Height, SWP_NOMOVE + SWP_NOSIZE)
End Function
Public Sub TimedMsgBox(strMessage As String, intSeconds As Integer)
    Dialog.Show
    Dialog.MsgDuration = intSeconds
    Dialog.Label1.Caption = strMessage
    SetAlwaysOnTop Dialog
    DoEvents
End Sub
            

Public Sub setpopup(nam As String, ord As String)
If DOPOPUP <> "1" Then
    Exit Sub
End If
Open "C:\WTRACK" For Output As #1
Print #1, "1"
NW$ = ""
Print #1, "2"
PTHW = nww
Print #1, "3"
PTHC = nwc
Print #1, "4"
PTHJ = nwj
Print #1, "5"
Close #1
If UCase(ord) = "F" Then
    Call CHANGELF(CStr(nam))
    NAMF = nam
Else
    Call changefl(CStr(nam))
    NAML = nam
End If
Dim db As Database, ds As Recordset
Dim addtopopup As Boolean
msg1 = ""
msg2 = ""
msg3 = ""
addtopopup = False
If PTHW > "" Then
    Set db = OpenDatabase(PTHW + "warrant.mdb")
    Set ds = db.OpenRecordset("select WARRANT, RECALLDATE, SENTON FROM warrantinfo where wname = " + Chr$(34) + NAML + Chr$(34) + " AND RECALLDATE IS NULL AND SENTON IS NULL")
    If Not ds.EOF Then
        ds.MoveFirst
        msg1 = "   Open Warrants: " + ds("warrant")
        ds.MoveNext
        While Not ds.EOF
            msg1 = msg1 + ", " + ds("warrant")
            ds.MoveNext
        Wend
        addtopopup = True
    End If
End If
If PTHC > "" Then
    Set db = OpenDatabase(PTHC + "civil.mdb")
    Set ds = db.OpenRecordset("select datereceived from magistrate where serviceof = " + Chr$(34) + NAMF + Chr$(34) + " AND served = '0' and nonservice = '0'")
    If Not ds.EOF Then
        ds.MoveFirst
        msg2 = "   Oustanding Civil Papers: Magistrate" + CStr(ds("datereceived"))
        ds.MoveNext
        While Not ds.EOF
            msg2 = msg2 + ", " + CStr(ds("datereceived"))
            ds.MoveNext
        Wend
        addtopopup = True
    End If
    Set ds = db.OpenRecordset("select datereceived from FamilyCourt where serviceof = " + Chr$(34) + NAMF + Chr$(34) + " AND served = '0' and nonservice = '0'")
    If Not ds.EOF Then
        ds.MoveFirst
        msg2 = "   Oustanding Civil Papers: FamilyCourt" + CStr(ds("datereceived"))
        ds.MoveNext
        While Not ds.EOF
            msg2 = msg2 + ", " + CStr(ds("datereceived"))
            ds.MoveNext
        Wend
        addtopopup = True
    End If
    Set ds = db.OpenRecordset("select datereceived from Executions where serviceof = " + Chr$(34) + NAMF + Chr$(34) + " AND served = '0' and nonservice = '0'")
    If Not ds.EOF Then
        ds.MoveFirst
        msg2 = "   Oustanding Civil Papers: Executions" + CStr(ds("datereceived"))
        ds.MoveNext
        While Not ds.EOF
            msg2 = msg2 + ", " + CStr(ds("datereceived"))
            ds.MoveNext
        Wend
        addtopopup = True
    End If
    Set ds = db.OpenRecordset("select datereceived from WritOther where serviceof = " + Chr$(34) + NAMF + Chr$(34) + " AND served = '0' and nonservice = '0'")
    If Not ds.EOF Then
        ds.MoveFirst
        msg2 = "   Oustanding Civil Papers: WritOther" + CStr(ds("datereceived"))
        ds.MoveNext
        While Not ds.EOF
            msg2 = msg2 + ", " + CStr(ds("datereceived"))
            ds.MoveNext
        Wend
        addtopopup = True
    End If
End If
If PTHJ > "" Then
    Set db = OpenDatabase(PTHJ + "jailsuite.mdb")
    Set ds = db.OpenRecordset("select casenumber from booking where sname = " + Chr$(34) + NAML + Chr$(34) + " AND releasedate is null")
    If Not ds.EOF Then
        ds.MoveFirst
        msg2 = "   Currently In Jail: " + CStr(ds("casenumber"))
        ds.MoveNext
        While Not ds.EOF
            msg2 = msg2 + ", " + CStr(ds("casenumber"))
            ds.MoveNext
        Wend
        addtopopup = True
    End If
End If
If addtopopup Then
    db.Close
    popup.Caption = "Alert Message - " + nam
    popup.msg1 = msg1
    popup.msg2 = msg2
    popup.msg3 = msg3
    popup.Show
End If
End Sub
Public Sub changefl(clf As String)
    hoLdname = clf
    osort1$ = ""
    If Left$(hoLdname, 1) = " " Then
        hoLdname = Mid$(hoLdname, 2)
    End If
    If InStr(hoLdname, " CORP") > 0 Or InStr(hoLdname, ",INC") > 0 Or InStr(hoLdname, "COMPANY") > 0 Or InStr(hoLdname, "INC.") > 0 Then
        osort1$ = hoLdname
    End If
    tso$ = hoLdname
    If InStr(tso$, " et al") > 0 Then
        tso$ = Left$(tso$, InStr(tso$, " et al") - 1)
    End If
    If InStr(tso$, " et. al.") > 0 Then
        tso$ = Left$(tso$, InStr(tso$, " et. al.") - 1)
    End If
    If InStr(tso$, ",et al") > 0 Then
        tso$ = Left$(tso$, InStr(tso$, ",et al") - 1)
    End If
    If InStr(tso$, ",et. al.") > 0 Then
        tso$ = Left$(tso$, InStr(tso$, ",et. al.") - 1)
    End If
    If Right$(tso$, 1) = "," Then
        tso$ = Left$(tso$, Len(tso$) - 1)
    End If
    If InStr(tso$, "&") > 0 Then
        tso$ = Left$(tso$, InStr(tso$, "&") - 1)
    End If
    If Right$(tso$, 1) = "," Then
        tso$ = Left$(tso$, Len(tso$) - 1)
    End If
    firstspace% = 0
    While Right$(tso$, 1) = " " And Len(tso$) > 1
        tso$ = Left$(tso$, Len(tso$) - 1)
    Wend
    For tt% = 1 To Len(tso$)
        If Mid$(tso$, tt%, 1) = "," Then
            firstspace% = tt%
            tt% = Len(tso$)
        End If
    Next tt%
    If firstspace% = 0 Then
        If osort1$ = "" Then
            osort1$ = tso$
        End If
        NAMF = osort1$
        Exit Sub
        'GoTo rsupdate
    End If
    tempsort$ = Mid$(tso$, firstspace% + 1)
    If Left$(tempsort$, 1) = " " Then
        tempsort$ = Mid$(tempsort$, 2)
    End If
    tso$ = Left$(tso$, firstspace% - 1)
    If Right$(tso$, 1) = " " Then
        tso$ = Left$(tso$, Len(tso$) - 1)
    End If
    tempsort$ = tempsort$ + " " + tso$
    If osort1$ = "" Then
        osort1$ = tempsort$
    End If
    If InStr(osort1$, "JR.") Then
        If Mid$(osort1$, InStr(osort1$, "JR.") + 3, 1) = " " Then
            osort1$ = Left$(osort1$, InStr(osort1$, "JR.") - 1) + Mid$(osort1$, InStr(osort1$, "JR.") + 4) + ", JR."
        Else
            osort1$ = Left$(osort1$, InStr(osort1$, "JR.") - 1) + Mid$(osort1$, InStr(osort1$, "JR.") + 3) + ", JR."
    End If
    End If
    If InStr(osort1$, "SR.") Then
        If Mid$(osort1$, InStr(osort1$, "SR.") + 3, 1) = " " Then
            osort1$ = Left$(osort1$, InStr(osort1$, "SR.") - 1) + Mid$(osort1$, InStr(osort1$, "SR.") + 4) + ", SR."
        Else
            osort1$ = Left$(osort1$, InStr(osort1$, "SR.") - 1) + Mid$(osort1$, InStr(osort1$, "SR.") + 3) + ", SR."
    End If
    End If
    If InStr(osort1$, "III") Then
        If Mid$(osort1$, InStr(osort1$, "III") + 3, 1) = " " Then
            osort1$ = Left$(osort1$, InStr(osort1$, "III") - 1) + Mid$(osort1$, InStr(osort1$, "III") + 4) + ", III"
        Else
            osort1$ = Left$(osort1$, InStr(osort1$, "III") - 1) + Mid$(osort1$, InStr(osort1$, "III") + 3) + ", III"
        End If
    End If
    If InStr(osort1$, "IV") Then
        If Mid$(osort1$, InStr(osort1$, "IV") + 2, 1) = " " Then
            osort1$ = Left$(osort1$, InStr(osort1$, "IV") - 1) + Mid$(osort1$, InStr(osort1$, "IV") + 3) + ", III"
        Else
            osort1$ = Left$(osort1$, InStr(osort1$, "IV") - 1) + Mid$(osort1$, InStr(osort1$, "IV") + 2) + ", III"
        End If
    End If
    If Left$(osort1$, 1) = " " Then
        osort1$ = Mid$(osort1$, 2)
    End If
    NAMF = osort1$
End Sub
Public Sub CHANGELF(cfl As String)
hoLdname = cfl
osort1$ = ""
If Left$(hoLdname, 1) = " " Then
    hoLdname = Mid$(hoLdname, 2)
End If
If InStr(hoLdname, " CORP") > 0 Or InStr(hoLdname, ",INC") > 0 Or InStr(hoLdname, "COMPANY") > 0 Or InStr(hoLdname, "INC.") > 0 Then
    osort1$ = hoLdname
End If
tso$ = hoLdname
If InStr(tso$, " et al") > 0 Then
    tso$ = Left$(tso$, InStr(tso$, " et al") - 1)
End If
If InStr(tso$, " et. al.") > 0 Then
    tso$ = Left$(tso$, InStr(tso$, " et. al.") - 1)
End If
If InStr(tso$, ",et al") > 0 Then
    tso$ = Left$(tso$, InStr(tso$, ",et al") - 1)
End If
If InStr(tso$, ",et. al.") > 0 Then
    tso$ = Left$(tso$, InStr(tso$, ",et. al.") - 1)
End If
If Right$(tso$, 1) = "," Then
    tso$ = Left$(tso$, Len(tso$) - 1)
End If
If InStr(tso$, "&") > 0 Then
    tso$ = Left$(tso$, InStr(tso$, "&") - 1)
End If
If Right$(tso$, 1) = " " Then
    tso$ = Left$(tso$, Len(tso$) - 1)
End If
firstspace% = 0
While Right$(tso$, 1) = " " And Len(tso$) > 1
    tso$ = Left$(tso$, Len(tso$) - 1)
Wend
For tt% = Len(tso$) To 1 Step -1
    If Mid$(tso$, tt%, 1) = " " Then
        If Mid$(tso$, tt% + 1, 3) = "JR." Or Mid$(tso$, tt% + 1, 3) = "SR." Or Mid$(tso$, tt% + 1, 3) = "III" Or Mid$(tso$, tt% + 1, 2) = "IV" Then
            aa = 1
        Else
            firstspace% = tt%
            tt% = 1
        End If
    End If
Next tt%
If firstspace% = 0 Then
    If osort1$ = "" Then
        osort1$ = tso$
    End If
    GoTo GGO
End If
tempsort$ = Mid$(tso$, firstspace% + 1)
If Left$(tempsort$, 1) = " " Then
    tempsort$ = Mid$(tempsort$, 2)
End If
tso$ = Left$(tso$, firstspace% - 1)
If Right$(tso$, 1) = " " Then
    tso$ = Left$(tso$, Len(tso$) - 1)
End If
tempsort$ = tempsort$ + ", " + tso$
If osort1$ = "" Then
    osort1$ = tempsort$
End If
'If InStr(osort1$, "JR.") Then
'    If Mid$(osort1$, InStr(osort1$, "JR.") + 3, 1) = " " Then
'        osort1$ = Left$(osort1$, InStr(osort1$, "JR.") - 1) + Mid$(osort1$, InStr(osort1$, "JR.") + 4) + ", JR."
'    Else
'        osort1$ = Left$(osort1$, InStr(osort1$, "JR.") - 1) + Mid$(osort1$, InStr(osort1$, "JR.") + 3) + ", JR."
'End If
'End If
'If InStr(osort1$, "SR.") Then
'    If Mid$(osort1$, InStr(osort1$, "SR.") + 3, 1) = " " Then
'        osort1$ = Left$(osort1$, InStr(osort1$, "SR.") - 1) + Mid$(osort1$, InStr(osort1$, "SR.") + 4) + ", SR."
'    Else
'        osort1$ = Left$(osort1$, InStr(osort1$, "SR.") - 1) + Mid$(osort1$, InStr(osort1$, "SR.") + 3) + ", SR."
'End If
'End If
'If InStr(osort1$, "III") Then
'    If Mid$(osort1$, InStr(osort1$, "III") + 3, 1) = " " Then
'        osort1$ = Left$(osort1$, InStr(osort1$, "III") - 1) + Mid$(osort1$, InStr(osort1$, "III") + 4) + ", III"
'    Else
'        osort1$ = Left$(osort1$, InStr(osort1$, "III") - 1) + Mid$(osort1$, InStr(osort1$, "III") + 3) + ", III"
'    End If
'End If
'If InStr(osort1$, "IV") Then
'    If Mid$(osort1$, InStr(osort1$, "IV") + 2, 1) = " " Then
'        osort1$ = Left$(osort1$, InStr(osort1$, "IV") - 1) + Mid$(osort1$, InStr(osort1$, "IV") + 3) + ", III"
'    Else
'        osort1$ = Left$(osort1$, InStr(osort1$, "IV") - 1) + Mid$(osort1$, InStr(osort1$, "IV") + 2) + ", III"
'    End If
'End If
If Left$(osort1$, 1) = " " Then
    osort1$ = Mid$(osort1$, 2)
End If
GGO:
NAML = osort1$

End Sub

Public Sub checkjailup(fromform As String)
foundjail = False
For t% = 0 To Forms.Count - 1
    Select Case LCase(Forms(t%).Name)
        Case "jsetup", "frmbookingreport", "frmclassificationsetup", "frmclassify", "frmdischarge", "frmdocarchive", "frminmateaffairs", "frminmateincident", "frmpropprofile", "frmreports", "frmstatexfer"
            If LCase(Forms(t%).Name) <> LCase(fromform) Then
                foundjail = True
            End If
    End Select
Next t%
If Not foundjail Then
    mainform.mdet.Caption = "Detention Center"
    For tb = 18 To 27
        mainform.Toolbar1.Buttons(tb).Visible = False
    Next tb
Else
    mainform.mdet.Caption = "Close Detention Center"
    For tb = 18 To 27
        mainform.Toolbar1.Buttons(tb).Visible = True
    Next tb
End If

End Sub
