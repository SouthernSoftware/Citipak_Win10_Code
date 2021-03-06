Attribute VB_Name = "modGLOpenFiles"
Option Explicit

Dim GLSetup   As GLSetupRecType
Dim GLFund    As GLFundRecType
Dim GLAcct    As GLAcctRecType
Dim GLFundIdx As GLFundIndexType
Dim GLAcctidx As GLAcctIndexType
Dim GLDept    As GLDeptRecType
Dim GLDeptIdx As GLDeptIndexType
Dim GLBank    As GLBankRecType
Dim APInvTax  As APInvTaxRecType
Dim GJEdit    As TrEditRecType
Dim GLTrans   As GLTransRecType
Dim CJEdit    As CJEditRecType
Dim BgtEdit   As TrEditRecType
Dim BgtTrans  As GLTransRecType
Dim OSChk     As OSChkRecType
Dim ApLedger  As APLedger81RecType
Dim APDist    As APDistRecType
Dim apvendor  As VendorRecType
Dim VendorIdx As VendorIdxRecType
Dim APVDist   As VendorDefDistRecType
Dim POControl As POControlRecType
Dim APPED     As POFORMRecType2
Dim POTrans   As GLTransRecType


'Public GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
Public Sub OpenSetupFile(SetupFileNum)
  Dim GLSetupRecLen  As Integer
  GLSetupRecLen = Len(GLSetup)
  SetupFileNum = FreeFile
  Open "GLSetup.DAT" For Random Shared As SetupFileNum Len = GLSetupRecLen
End Sub
Public Sub OpenFundFile(FundFileNum, NumFundRecs As Integer)
  Dim GLFundRecLen As Integer
  GLFundRecLen = Len(GLFund)
  FundFileNum = FreeFile
  Open "GLFund.DAT" For Random Shared As FundFileNum Len = GLFundRecLen
  NumFundRecs = LOF(FundFileNum) \ GLFundRecLen
End Sub

Public Sub OpenFundIdx(FundIdxFileNum, NumFIdxRecs)
  Dim GLFundIdxLen As Integer
  GLFundIdxLen = Len(GLFundIdx)
  FundIdxFileNum = FreeFile
  Open "GLFund.Idx" For Random Shared As FundIdxFileNum Len = GLFundIdxLen
  NumFIdxRecs = LOF(FundIdxFileNum) \ GLFundIdxLen
End Sub
Public Sub OpenAcctFile(AcctFileNum, Optional NumAccts As Integer)
  Dim GLAcctRecLen As Integer
  GLAcctRecLen = Len(GLAcct)
  AcctFileNum = FreeFile
  Open "GLAcct.DAT" For Random Shared As AcctFileNum Len = GLAcctRecLen
  NumAccts = LOF(AcctFileNum) \ GLAcctRecLen
End Sub

Public Sub OpenAcctIdx(AcctIdxFileNum, NumAIdxRecs)
  Dim GLAcctIdxLen As Integer
  GLAcctIdxLen = Len(GLAcctidx)
  AcctIdxFileNum = FreeFile
  Open "GLAcct.Idx" For Random Shared As AcctIdxFileNum Len = GLAcctIdxLen
  NumAIdxRecs = LOF(AcctIdxFileNum) \ GLAcctIdxLen
End Sub
Public Sub OpenDeptFile(DeptFileNum, NumDeptRecs As Integer)
  Dim GLDeptRecLen As Integer
  GLDeptRecLen = Len(GLDept)
  DeptFileNum = FreeFile
  Open "GLDept.DAT" For Random Shared As DeptFileNum Len = GLDeptRecLen
  NumDeptRecs = LOF(DeptFileNum) \ GLDeptRecLen
End Sub
Public Sub OpenDeptIdx(DeptIdxFileNum, NumDIdxRecs)
  Dim GLDeptIdxLen As Integer
  GLDeptIdxLen = Len(GLDeptIdx)
  DeptIdxFileNum = FreeFile
  Open "GLDept.Idx" For Random Shared As DeptIdxFileNum Len = GLDeptIdxLen
  NumDIdxRecs = LOF(DeptIdxFileNum) \ GLDeptIdxLen
End Sub
Public Sub OpenBankFile(BankFileNum, NumBankRecs As Integer)
  Dim GLBankRecLen As Integer
  GLBankRecLen = Len(GLBank)
  BankFileNum = FreeFile
  Open "GLBank.DAT" For Random Shared As BankFileNum Len = GLBankRecLen
  NumBankRecs = LOF(BankFileNum) \ (GLBankRecLen)
End Sub
Public Sub OpenInvTaxFile(InvTaxFileNum)
  Dim APInvTaxRecLen As Integer
  APInvTaxRecLen = Len(APInvTax)
  InvTaxFileNum = FreeFile
  Open "APInvTax.DAT" For Random Shared As InvTaxFileNum Len = APInvTaxRecLen
End Sub
Public Sub OpenGJEditFile(GJEditFileNum, NumEdTrans)
  On Local Error GoTo GJError
  Dim GJEdLen As Integer
  GJEdLen = Len(GJEdit)
  GJEditFileNum = FreeFile
  Open "GLGJED.DAT" For Random Shared As GJEditFileNum Len = GJEdLen
  NumEdTrans = LOF(GJEditFileNum) \ (GJEdLen)
  Lock #GJEditFileNum
  Exit Sub

GJError:
  GJEditFileNum = -1
  'Close BgtEditFileNum
  MsgBox "The General Journal File Has Been Opened By Another User, And May Not Be Accessed At This Time.", vbOKOnly, "Access Denied"

End Sub
Public Sub OpenTransFile(TransFileNum, NumTrans As Long)
  Dim TransRecLen As Integer
  TransRecLen = Len(GLTrans)
  TransFileNum = FreeFile
  Open "GLTRANS.DAT" For Random Shared As TransFileNum Len = TransRecLen
  NumTrans = LOF(TransFileNum) \ (TransRecLen)
End Sub
Public Sub OpenCJEditFile(CJEditFileNum, NumEdTrans, CJType)
  Dim CJEdLen As Integer
  CJEdLen = Len(CJEdit)
  CJEditFileNum = FreeFile
  
  Select Case CJType
  Case 1
    Open "GLCRED.DAT" For Random Shared As CJEditFileNum Len = CJEdLen
    NumEdTrans = LOF(CJEditFileNum) \ (CJEdLen)
  Case 2
    Open "GLCDED.DAT" For Random Shared As CJEditFileNum Len = CJEdLen
    NumEdTrans = LOF(CJEditFileNum) \ (CJEdLen)
  End Select

End Sub
'Public Sub OpenCJREditFile(CJEditFileNum, NumEdTrans)
'  Dim CJEdLen As Integer
'  CJEdLen = Len(GLCREd)
'  CJEditFileNum = FreeFile
'  Open "Glcred.dat" For Random Shared As CJEditFileNum Len = CJEdLen
'  NumEdTrans = LOF(CJEditFileNum) \ (CJEdLen)
'End Sub

Public Sub OpenOSChkFile(OSChkFileNum, NumOSChks)
  Dim OSChkLen As Integer
  OSChkLen = Len(OSChk)
  OSChkFileNum = FreeFile
  Open "crchk.dat" For Random Shared As OSChkFileNum Len = OSChkLen
  NumOSChks = LOF(OSChkFileNum) \ (OSChkLen)
End Sub
Public Sub OpenBgtTransFile(BgtTransFile, NumTrans)
  Dim BgtTransLen As Integer
  BgtTransLen = Len(BgtTrans)
  BgtTransFile = FreeFile
  Open "BGTTRANS.dat" For Random Shared As BgtTransFile Len = BgtTransLen
  NumTrans = LOF(BgtTransFile) \ BgtTransLen
End Sub
Public Sub OpenBgtEditFile(BgtEditFileNum, NumEdTrans)
  On Local Error GoTo BgtError
  Dim BgtEdLen As Integer
  BgtEdLen = Len(BgtEdit)
  BgtEditFileNum = FreeFile
  Open "BGTED.dat" For Random Shared As BgtEditFileNum Len = BgtEdLen
  NumEdTrans = LOF(BgtEditFileNum) \ BgtEdLen
  Lock #BgtEditFileNum
  Exit Sub

BgtError:
  BgtEditFileNum = -1
  'Close BgtEditFileNum
  MsgBox "The Budget Edit File Has Been Opened By Another User, And May Not Be Accessed At This Time.", vbOKOnly, "Access Denied"
End Sub
Public Sub OpenAPLedgerFile(APLedgerFile, NumTran&, RecLen)
  RecLen = Len(ApLedger)
  APLedgerFile = FreeFile
  Open "APLEDGER.DAT" For Random Shared As APLedgerFile Len = RecLen
  NumTran& = LOF(APLedgerFile) \ RecLen
End Sub
Public Sub OpenAPDistFile(APDistFile, NumDistRecs&, RecLen)
  RecLen = Len(APDist)
  APDistFile = FreeFile
  Open "APDIST.DAT" For Random Shared As APDistFile Len = RecLen
  NumDistRecs& = LOF(APDistFile) \ RecLen
End Sub
Public Sub OpenVendorFile(VendorFile, NumVRecs)
  Dim VRecLen As Integer
  VRecLen = Len(apvendor)
  VendorFile = FreeFile
  Open "apvendor.dat" For Random Shared As VendorFile Len = VRecLen
  NumVRecs = LOF(VendorFile) \ VRecLen
End Sub
Public Sub OpenVendorIdx(VendorIdxFile, NumActiveVendors)
  Dim VendorIdxLen As Integer
  VendorIdxLen = Len(VendorIdx)
  VendorIdxFile = FreeFile
  'OPEN "apvendor.idx" FOR RANDOM ACCESS READ WRITE SHARED AS #4 LEN = 14
  Open "apvendor.idx" For Random Shared As VendorIdxFile Len = VendorIdxLen
  NumActiveVendors = LOF(VendorIdxFile) \ VendorIdxLen
End Sub

Public Sub OpenDefDistFile(DefRecLen, APDefDistFile, NumDefRecs)
  DefRecLen = Len(APVDist)
  APDefDistFile = FreeFile
  Open "APVDIST.DAT" For Random Shared As APDefDistFile Len = DefRecLen
  NumDefRecs = LOF(APDefDistFile) \ DefRecLen
End Sub
Public Sub OpenPOFile(POFile, NumRecs)
  Dim POFileLen As Integer
  POFileLen = Len(POControl)
  POFile = FreeFile
  Open "APPOCRL.DAT" For Random Shared As #POFile Len = POFileLen
  NumRecs = LOF(POFile) \ POFileLen
End Sub
Public Sub OpenPOEditFile(POEditFile, NumEditRecs)
  Dim EdLen As Integer
  EdLen = Len(APPED)
  POEditFile = FreeFile
  Open "APPED.DAT" For Random Shared As POEditFile Len = EdLen
  NumEditRecs = LOF(POEditFile) \ (EdLen)
End Sub
Public Sub OpenPOTransFile(TransFileNum, NumTrans&)
  Dim TransRecLen As Integer
  ReDim TempPOTrans(1) As GLTransRecType
  TransRecLen = Len(TempPOTrans(1))
  TransFileNum = FreeFile
  Open "POTRANS.DAT" For Random Shared As TransFileNum Len = TransRecLen
  NumTrans& = LOF(TransFileNum) \ TransRecLen
  Erase TempPOTrans

End Sub

