Attribute VB_Name = "ExportStartup"
Option Explicit
DefInt A-Z

Sub Main()
  ImpExpUSenHHInfo True
  SmallPause 2
  End
End Sub


'************************************************************************************
Private Sub ImpExpUSenHHInfo(ByVal ImpExpFlag As Boolean)

  Dim HighRead As Double, ILowRead As Double
  Dim IdxRecLen As Integer, IdxNumOfRecs As Integer, IdxFileSize As Long
  Dim UBFile As Integer
  Dim cnt As Long, CustName As String, CustRec As Long
  Dim HighReadPerc As Double
  Dim Prec As Long, RecNumber As Long
  Dim PrevRead As String, HighReadS As String
  Dim LowRead As Double, LowReadS As String, FileHdrLine As String
  Dim outfile As Integer
  Dim PrevDate$ ', Numsent As Long
  Dim Mout As String
  Dim Chkbook As Boolean, RType As String, MPIDNO As String, metercnt As Integer
  Dim Prev#
  Dim UBCustRec(1) As NewUBCustRecType
  Dim CustRecNo As Long, UBCustRecLen As Integer
  Dim cc$
  
  Dim UBSetUpRec(1) As UBSetupRecType
  Dim UBSetupLen As Integer
  
  UBSetupLen = Len(UBSetUpRec(1))
  Open "UBSETUP.DAT" For Random Shared As #10 Len = UBSetupLen    'open data file
  Get #10, 1, UBSetUpRec(1)
  Close #10
  TOWNNAME$ = QPTrim$(UBSetUpRec(1).UTILNAME)
  
  UBCustRecLen = Len(UBCustRec(1))

  Dim abc$ ', zz%, xx%, yy%, gg$
  '67,65,78,68,79,82 CANDOR
  abc$ = Chr$(67) + Chr$(65) + Chr$(78) + Chr$(68) + Chr$(79) + Chr$(82)

  If InStr(TOWNNAME, abc) <= 0 Then
    End
  End If
  
  IdxRecLen = 4               'we are using an integer
  IdxFileSize& = FileSize&("UBCUSTBK.IDX")
  IdxNumOfRecs = IdxFileSize& \ IdxRecLen
  ReDim IdxBuff(1 To IdxNumOfRecs) As UBCustIndexRecType
  FGetAH "UBCUSTBK.IDX", IdxBuff(), IdxRecLen, IdxNumOfRecs  'load it
  

SendReads:
  cc$ = ","
  ReDim Mtr(1) As UBPMHHSendRecType
  UBCustRecLen = Len(UBCustRec(1))
  UBFile = FreeFile
  Open "UBCUST.DAT" For Random Shared As UBFile Len = UBCustRecLen
  
'  IdxNumOfRecs = LOF(UBFile) \ UBCustRecLen
'  For cnt = 1 To IdxNumOfRecs
'    Get UBFile, cnt, UBCustRec(1)
'    UBCustRec(1).SEQNUMB = cnt
'    UBCustRec(1).BILLCYCL = 1
'    Put UBFile, cnt, UBCustRec(1)
'  Next
'  Close
'  End
  
  outfile = FreeFile
  Open "DataSync.CSV" For Output As #outfile
  
  FrmShowPctComp.Label1 = "Exporting Reading Information."
  FrmShowPctComp.CmdCancel.Enabled = False
  FrmShowPctComp.AutoClose = "0"
  FrmShowPctComp.Show '1, Parent
  DoEvents

  SmallPause 2
'-----------------------------------------------------
  For cnt = 1 To IdxNumOfRecs
    Prec& = IdxBuff(cnt).RecNum
    RecNumber = Prec&
      If Not (Prec&) = 0 Then
        Get UBFile, Prec&, UBCustRec(1)
        FrmShowPctComp.ShowPctComp cnt, IdxNumOfRecs
        If (UBCustRec(1).Status <> "F") And Val(UBCustRec(1).Book) > 0 Then
          If UBCustRec(1).DelFlag = 0 Then
          For metercnt = 1 To 7
            If Len(QPTrim$(UBCustRec(1).LocMeters(metercnt).MtrNum)) > 0 Then
              Mtr(1).RecordID = "M1"
              Mtr(1).AccountNum = LTrim$(Str$(RecNumber))
              Mtr(1).MeterNum = UBCustRec(1).LocMeters(metercnt).MtrNum
              Mtr(1).SeqNum = UBCustRec(1).SEQNUMB
              Mtr(1).ReadTypeCode = UBCustRec(1).LocMeters(metercnt).MTRType
              If UBCustRec(1).LocMeters(metercnt).CurRead > 0 Then
                Prev# = UBCustRec(1).LocMeters(metercnt).CurRead
              Else
                Prev# = 0
              End If
              PrevRead$ = LTrim$(Str$(Prev#))
              PrevRead$ = PrevRead$ + String$(10 - Len(PrevRead$), " ")
              If UBCustRec(1).LocMeters(metercnt).AvgUse < 0 Then
                UBCustRec(1).LocMeters(metercnt).AvgUse = 0
              End If
              HighRead# = Fix(((0# + UBCustRec(1).LocMeters(1).AvgUse) * (HighReadPerc / 100) + UBCustRec(1).LocMeters(1).CurRead))
              If HighRead# <= 1 Then
                HighRead# = 1
              End If
              HighReadS$ = LTrim$(Str$(HighRead#))
              If Len(HighReadS$) < 10 Then
                HighReadS$ = HighReadS$ + String$(10 - Len(HighReadS$), " ")
              End If
              LowRead# = PrevRead$ '1 'UBCustRec(1).LocMeters(1).CurRead
              LowReadS$ = LTrim$(Str$(LowRead#))
              If Len(LowReadS$) < 10 Then
                LowReadS$ = LowReadS$ + String$(10 - Len(LowReadS$), " ")
              End If
              PrevDate$ = Num2Date(UBCustRec(1).LocMeters(metercnt).CurDate)
              PrevDate$ = Left$(PrevDate$, 2) + Mid$(PrevDate$, 4, 2)
              Mtr(1).HighRead = HighReadS$
              Mtr(1).LowRead = LowReadS$
              Mtr(1).LocatnCode = String$(4, " ")
              Mtr(1).InstrCode = String$(4, " ")
              Mtr(1).AcctCode = UBCustRec(1).Status
              Mtr(1).Address = LTrim$(Left$(UBCustRec(1).ServAddr, 40))
              Mtr(1).Name = LTrim$(Left$(UBCustRec(1).CustName, 20))
              Mtr(1).PrevRead = PrevRead$
              Mtr(1).PReadDate = PrevDate$
              Mtr(1).NumDials = "0"
              Mtr(1).DecimalPL = " "
              Mtr(1).Compound = String$(10, " ")
              Mtr(1).UID = QPTrim$(UBCustRec(1).LocMeters(metercnt).MtrIDNO)
              Mtr(1).Longitude = UBCustRec(1).LocMeters(metercnt).MtrLng
              Mtr(1).Latitude = UBCustRec(1).LocMeters(metercnt).MtrLat
              Mtr(1).Filler = String$(10, " ")
              Mtr(1).CrLf = Chr$(13) + Chr$(10)
              
              Mout$ = QPTrim(Mtr(1).RecordID) + cc$ + QPTrim(Mtr(1).AccountNum) + cc$
              Mout$ = Mout$ + QPTrim(Mtr(1).MeterNum) + cc$ + QPTrim(Mtr(1).SeqNum) + cc$
              Mout$ = Mout$ + QPTrim(Mtr(1).ReadTypeCode) + cc$ + QPTrim(Mtr(1).HighRead) + cc$
              Mout$ = Mout$ + QPTrim(Mtr(1).LowRead) + cc$
              Mout$ = Mout$ + QPTrim(Mtr(1).LocatnCode) + cc$ + QPTrim(Mtr(1).InstrCode) + cc$
              Mout$ = Mout$ + QPTrim(Mtr(1).AcctCode) + cc$
              Mout$ = Mout$ + QPTrim(Mtr(1).Address) + cc$ + QPTrim(Mtr(1).Name) + cc$
              Mout$ = Mout$ + QPTrim(Mtr(1).PrevRead) + cc$
              Mout$ = Mout$ + QPTrim(Mtr(1).PReadDate) + cc$ + QPTrim(Mtr(1).NumDials) + cc$
              Mout$ = Mout$ + QPTrim(Mtr(1).DecimalPL) + cc$
              Mout$ = Mout$ + QPTrim(Mtr(1).Compound) + cc$ + QPTrim(Mtr(1).UID) + cc$
              Mout$ = Mout$ + QPTrim(Mtr(1).Longitude) + cc$
              Mout$ = Mout$ + QPTrim(Mtr(1).Latitude) + cc$ + QPTrim(Mtr(1).Filler) + cc$ + Mtr(1).CrLf
              Print #outfile, Mout$;
           End If
         Next metercnt
       End If
     End If
   End If
  Next
'-----------------------------------------------------
  Close
End Sub

Private Sub SmallPause(ByVal PauseAmt As Integer)
    Static st1!, st2!
    st1! = Timer
    st2! = st1! + PauseAmt
    Do Until Timer > st2!
    Loop
End Sub

Private Sub FGetAH(FileName As String, IdxBuff() As UBCustIndexRecType, ByVal IdxRecLen As Integer, ByVal IdxNumOfRecs As Long)
  Dim ICnt As Long
  Dim IdxFile As Integer
  IdxFile = FreeFile
  Open FileName For Random Shared As IdxFile Len = IdxRecLen
  For ICnt = 1 To IdxNumOfRecs
    Get IdxFile, ICnt, IdxBuff(ICnt).RecNum
  Next
  Close IdxFile
End Sub

Public Function QPTrim$(Text As String)
  Dim StrLen As Long
  Dim cnt As Long
  Dim thischar As Integer
  StrLen = Len(Text)
  For cnt = 1 To StrLen
    thischar = Asc(Mid$(Text, cnt, 1))
    If thischar = 0 Or thischar = 44 Then
      Mid$(Text$, cnt, 1) = " "
    End If
  Next
  QPTrim$ = Trim$(Text)
End Function

Public Function Num2Date$(intDate%)
  On Error GoTo BadNum2Date
  If intDate% = -32767 Then
    Num2Date$ = ""
  Else
    Num2Date$ = Format(DateAdd("d", (intDate%), "12-31-1979"), "mm/dd/yyyy")
  End If
  Exit Function
BadNum2Date:
  On Error GoTo 0
  Num2Date = ""
End Function

Public Function Date2Num%(txtDate$)
  On Error GoTo BadDate2Num
  If Len(QPTrim$(txtDate$)) = 10 Then
    Date2Num% = DateDiff("d", "12/31/1979", txtDate$)
  Else
    Date2Num% = -32767
  End If
  Exit Function

BadDate2Num:
  On Error GoTo 0
  Date2Num% = -32767
End Function
Public Sub KillFileD(FileName$)
  On Local Error GoTo ErrorCatch
  If ExistD(FileName$) Then
    Kill FileName$
  End If
  Exit Sub
  
ErrorCatch:
  Select Case Err
    Case Is <> 53
      MsgBox ("File deletion permission denied " + Str$(Err) + " . PLEASE CONTACT SOUTHERN SOFTWARE @ 1-800-842-8190."), vbOKOnly
    Case 53
      Resume ExitFillFile
  End Select
    
ExitFillFile:
  
End Sub
Public Function ExistD(FileName$)
  On Local Error Resume Next
  Dim FileHandle As Integer
  Dim FileSize As Long
  FileHandle = FreeFile
  Open FileName$ For Binary As FileHandle
  FileSize = LOF(FileHandle)
  Close FileHandle
  
  If FileSize > 0 Then
    ExistD = True
  Else
    ExistD = False
    Kill FileName$
  End If
End Function

Public Function Exist(FileName$)
  On Local Error Resume Next
  Dim FileHandle As Integer
  Dim FileSize As Long
  FileHandle = FreeFile
  Open FileName$ For Binary As FileHandle
'  If Err Then
'    FileName$ = ""
'  End If
  FileSize = LOF(FileHandle)
  Close FileHandle
  
  If FileSize > 0 Then
    Exist = True
  Else
    Exist = False
   ' If UCase(FileName$) <> UCase(UBPath$ + "UBCust.Dat") Then  'Added this to see if corrects problem with custfile getting deleted. 10/05/2007 -PS

      'Kill FileName$
   ' End If
  End If
'  On Local Error GoTo 0
End Function

Public Sub KillFile(FileName$)
  On Local Error GoTo ErrorCatch
  If Exist(FileName$) Then
    Kill FileName$
  End If
  Exit Sub
  
ErrorCatch:
  Select Case Err
    Case Is <> 53
      MsgBox ("File deletion permission denied " + Str$(Err) + " . PLEASE CONTACT SOUTHERN SOFTWARE @ 1-800-842-8190."), vbOKOnly
    Case 53
      Resume ExitFillFile
  End Select
    
ExitFillFile:
  
End Sub

Public Function FileSize(FileName$) As Long
  Dim FileHandle As Integer
  If Exist(FileName$) Then
    FileHandle = FreeFile
    Open FileName$ For Binary As FileHandle
    FileSize = LOF(FileHandle)
    Close FileHandle
  Else
    FileSize = 0
    
  End If
End Function

