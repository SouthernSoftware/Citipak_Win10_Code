CONST PathData$ = "UBBAGPTH.DAT"

TYPE BADReadRecType
  Fill          AS STRING * 128
END TYPE

TYPE BADReadRecType0
 RecordID       AS STRING * 3
'***
 Pad1           AS STRING * 11 '14
 Account        AS STRING * 20 '34
'***
 CurRead        AS STRING * 10 '44
 pad2           AS STRING * 3  '47
 ReadDate       AS STRING * 8  '56
 Pad            AS STRING * 73 '
END TYPE

TYPE FDRTYPE
 RecordID       AS STRING * 3
 TABLE          AS STRING * 1
 PROBE          AS STRING * 1
 VERS           AS STRING * 5
 Reserved       AS STRING * 3
 CYCLES         AS STRING * 2
 RADIOREADYN    AS STRING * 1
 WANDYN         AS STRING * 1
 EXTFORMATYN    AS STRING * 1
 Pad            AS STRING * 108
END TYPE

TYPE CDRTYPE
 RecordID       AS STRING * 3
 CycleNumber    AS STRING * 2
 NumberCycles   AS STRING * 4
 CycleDate      AS STRING * 8
 Pad            AS STRING * 109
END TYPE

TYPE RDRType
 RecordID        AS STRING * 3
 RouteNumber     AS STRING * 8
 SurveyYN        AS STRING * 1
 RouteMessYN     AS STRING * 1
 Keys            AS STRING * 4
 Readings        AS STRING * 4
 Demands         AS STRING * 4
 Keyed           AS STRING * 4
 Probed          AS STRING * 4
 Radio           AS STRING * 4
 Customers       AS STRING * 4
 Meters          AS STRING * 4
 TimeAllowed     AS STRING * 6
 Gas             AS STRING * 4
 Water           AS STRING * 4
 Electric        AS STRING * 4
 Location        AS STRING * 4
 Extra           AS STRING * 4
 Region          AS STRING * 2
 Zone            AS STRING * 2
 Office          AS STRING * 2
 BillCycle       AS STRING * 2
 DropCycle       AS STRING * 2
 WandReads       AS STRING * 4
 AMR             AS STRING * 1
 Pad             AS STRING * 40
END TYPE

TYPE CUSType
 RecordID       AS STRING * 3
 RouteNumber    AS STRING * 8
 Meters         AS STRING * 3
 AccountNumber  AS STRING * 20
 Name           AS STRING * 20
 Address1       AS STRING * 20
 Address2       AS STRING * 20
 Reserved       AS STRING * 2
 Group          AS STRING * 1
 CustInfo       AS STRING * 20
 ExCustRec      AS STRING * 1
 Segment        AS STRING * 4
 UtilityID      AS STRING * 2
 PassThur       AS STRING * 1
 Pad            AS STRING * 1
END TYPE

TYPE MTRType
 RecordID       AS STRING * 3   '3
 RouteNumber    AS STRING * 8   '11
 Reads          AS STRING * 3   '14
 Reserved       AS STRING * 2   '16
 Group          AS STRING * 1   '17
 Reserved1      AS STRING * 1   '18
 Reserved2      AS STRING * 7   '25
 Survey         AS STRING * 1   '26
 Reserved3      AS STRING * 2   '28
 Survey2        AS STRING * 1   '29
 BillCode       AS STRING * 1   '30
 MtrStatus      AS STRING * 1   '31
 OpticalProbe   AS STRING * 14  '45
 MtrNumber      AS STRING * 12  '57
 Reserved4      AS STRING * 2   '59
 MeterType      AS STRING * 2   '61
 MeterSeq       AS STRING * 8   '69
 MeterInfo      AS STRING * 20  '89
 Reserved5      AS STRING * 1   '90
 Location       AS STRING * 2   '92
 Reserved6      AS STRING * 1   '93
 ReadInst1      AS STRING * 2   '95
 Reserved7      AS STRING * 1   '96
 ReadInst2      AS STRING * 2   '98
 SpecMessage    AS STRING * 1   '99
 Reserved8      AS STRING * 1   '100
 SpecMessageYN  AS STRING * 1   '101
 MtrCat         AS STRING * 1   '102
 LocExtraMtr    AS STRING * 1   '103
 TimeCode       AS STRING * 3   '106
 MtrAudit1      AS STRING * 2   '108
 MtrAudit2      AS STRING * 2   '110
 MtrAudit3      AS STRING * 2   '112
 MtrAudit4      AS STRING * 2   '114
 Pad            AS STRING * 12  '126
END TYPE

TYPE RDGType
 RecordID       AS STRING * 3
 RouteNumber    AS STRING * 8  '11
 Text           AS STRING * 4  '15
 Prompt         AS STRING * 1  '16
 ReadDir        AS STRING * 1  '17
 Compare        AS STRING * 3  '20
 Validation     AS STRING * 3  '23
 Reserved       AS STRING * 1  '24
 Channel        AS STRING * 2  '26
 Dials          AS STRING * 2  '28
 Decimals       AS STRING * 2  '30
 ReadMethod     AS STRING * 1  '31
 PrevRead       AS STRING * 10 '41
 High1          AS STRING * 10 '51
 Low1           AS STRING * 10 '61
 MtrConstant    AS STRING * 6  '67
 ConstantFlag   AS STRING * 1  '68
 HHFFlag        AS STRING * 1  '69
 PosCreep       AS STRING * 5  '74
 Estimates      AS STRING * 1  '75
 Reserved2      AS STRING * 1  '76
 ReadType       AS STRING * 2  '78
 MaxPercent     AS STRING * 6  '84
 MinPercent     AS STRING * 6  '
 NegCreep       AS STRING * 5  '95
 Pad            AS STRING * 31 '
END TYPE

TYPE RFFType
 RecordID       AS STRING * 3
 RouteNumber    AS STRING * 8
 RFERT          AS STRING * 8
 Reserved       AS STRING * 6
 RFProgram      AS STRING * 4
 Reserved2      AS STRING * 7
 GeoArea        AS STRING * 2
 RFFreq         AS STRING * 12
 RFTone         AS STRING * 4
 Reserved3      AS STRING * 10
 Tamper         AS STRING * 2
 ConcIndicator  AS STRING * 1
 Pad            AS STRING * 59
END TYPE

TYPE WRRType
 RecordID       AS STRING * 3
 RouteNumber    AS STRING * 8
 DeviceID       AS STRING * 14
 WandProg       AS STRING * 4
 Resv1          AS STRING * 5
 Resv2          AS STRING * 1
 NodeNumb       AS STRING * 2
 NoTamper       AS STRING * 1  'y/n "Y" = no Check
 Fill1          AS STRING * 88
END TYPE

