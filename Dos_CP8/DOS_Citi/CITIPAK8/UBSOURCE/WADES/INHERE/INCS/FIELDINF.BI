'----- Define the Field Information TYPE
TYPE FieldInfo
     Fields    AS INTEGER              '0
     Row       AS INTEGER              '1
     LCol      AS INTEGER              '2
     RCol      AS INTEGER              '3
     StorLen   AS INTEGER              '4
     FType     AS INTEGER              '5
     RelFile   AS STRING * 8           '6
     RelFld    AS INTEGER              '10
     Indexed   AS INTEGER              '11
     FldName   AS STRING * 8           '12
     Decimals  AS INTEGER              '16
     RelHandle AS INTEGER              '17
     Protected AS INTEGER              '18
     ScratchI  AS INTEGER
     LowRange  AS DOUBLE
     HiRange   AS DOUBLE
     ScratchS  AS STRING * 8
END TYPE

CONST FrmWdth = 64
CONST MaxFields = 255

'----- Define Field Type constants
CONST StrFld = 1
CONST PropStrFld = 23
CONST UCaseStrFld = 2
CONST NumericStrFld = 16
CONST NotesFld = 17
CONST IntFld = 3
CONST LongIntFld = 18
CONST SngFld = 4
CONST DblFld = 5
CONST DateFld = 6
CONST EuroDateFld = 19
CONST PhoneFld = 7
CONST SoSecFld = 8
CONST ZipFld = 9
CONST MoneyFld = 10
CONST VirtualFld = 12
CONST Relational = 11
CONST MultChFld = 13
CONST MultChAFld = 22
CONST LogicalFld = 14
CONST RecNoFld = 20
CONST TotRecsFld = 21
CONST SequFld = 15
CONST Button = 24

CONST F1Key = -59
CONST F2Key = -60
CONST F3Key = -61
CONST F4KEY = -62
CONST F5KEY = -63
CONST F6KEY = -64
CONST F7KEY = -65
CONST F8KEY = -66
CONST F9KEY = -67
CONST F0Key = -68
CONST F10Key = -68
CONST HomeKey = -71
CONST EndKey = -79
CONST TabKey = 9
CONST ShftTabKey = -15
CONST CtrlR = -116
CONST CtrlL = -115

CONST PgUpKey = -73
CONST PgDnKey = -81
CONST ESC = 27
CONST EscKey = 27
CONST EnterKey = 13
CONST SKey = 83
CONST PKey = 80

CONST MainMenu = 1
CONST EmpMaintMenu = 2
CONST PayProcessMenu = 3
CONST ManMenu = 4
CONST YTDMenu = 5
CONST ReportsMenu = 6
CONST ControlMenu = 7
CONST DraftMenu = 10

CONST UseY = 89
CONST UseN = 78

'KeyCodes


