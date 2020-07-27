DEFINT A-Z
DECLARE SUB InitializeSystem ()
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)

'$INCLUDE: 'DefCnf.bi'                  'Defines TYPE for monitor/color info.
'$INCLUDE: 'FieldInf.bi'                '        "        field information
'$INCLUDE: 'FormEdit.bi'                '        "        form editing info.

'DIM Statements go here

InitializeSystem

'your code start here:

'example

  PrintRptFile "Print AUTOEXEC.BAT File!", "c:\autoexec.bat", 1, RetCode%, 1



