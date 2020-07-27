'Include file for declaring procedures used by QuickScreen

'Procedure from VERTMENU.BAS
DECLARE SUB VertMenu (Item$(), Choice%, MaxLen%, BoxBot%, Ky$, Action%, Cnf AS ANY)

'Procedures from FORMLIB.BAS
DECLARE SUB LibGetFldDef (LibName$, ScrName$, StartEl, Fld() AS FieldInfo, Form$(), ErrCode)
DECLARE FUNCTION LibNumberOfFields% (LibName$, ScrName$)

'Procedures from FORMEDIT.BAS
DECLARE SUB CalcFields (Start, FldNo%, Form$(), Fld() AS ANY)
DECLARE SUB EditForm (Form$(), Fld() AS FieldInfo, Frm AS FormInfo, Cnf AS ANY, Action%)
DECLARE FUNCTION FldNum% (FldName$, Fld() AS FieldInfo)
DECLARE SUB PrintArray (FirstFld%, LastFld%, Form$(), Fld() AS FieldInfo)
DECLARE SUB SaveField (FldNo%, Form$(), Fld() AS ANY, BadFld%)
DECLARE SUB UnPackBuffer (FirstFld%, LastFld%, Form$(), Fld() AS FieldInfo)
DECLARE FUNCTION Value# (E$, ErrCode%)

'Procedure from SCRNLIB.BAS
DECLARE SUB LibFile2Scrn (LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%)

