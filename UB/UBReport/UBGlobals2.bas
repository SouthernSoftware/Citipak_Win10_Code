Attribute VB_Name = "UBGlobals2"
Option Explicit

Type UBDGProcRecType        ' File Layout for Sending Out Records
    RouteID As String * 20
    SvcTyp As String * 1
    CustName As String * 25
    SvcLoc As String * 21
    MeterSN As String * 20
    MeterType As String * 1       ' C for reg mtr   D for demand elec
    High As String * 10
    Low As String * 10
    Msg As String * 110
    Account As String * 10
    NewRdng As String * 10
    NewDmnd As String * 10
    Date As String * 6
    Time As String * 6
    NewAcctRte As String * 20
End Type
  
Type UBBadgerBeaconRecType
'Account
  AcctID          As String * 32
  AcctFName       As String * 64
  AcctLName       As String * 64
  AcctFullName    As String * 128
  AcctEMail       As String * 254
  AcctPhone       As String * 32
  AcctPhoneExt    As String * 255
  AcctAddr1       As String * 64
  AcctAddr2       As String * 64
  AcctAddr3       As String * 64
  AcctCity        As String * 64
  AcctState       As String * 2
  AcctZip         As String * 10
  AcctCountry     As String * 3
  AcctPerID       As String * 32
  AcctStatus      As String * 1
  AcctPortStat    As String * 1
  AcctBLCyc       As String * 12   'not Citipak's billing cycle
  AcctPaperless   As String * 1
  AcctAutoPay     As String * 1
  AcctBillerAP    As String * 1

'Location
  LocID           As String * 40   'NA
  LocName         As String * 64   'Name of location CP has no location name NA
  LocAddParity    As String * 1    'Odd or even house number NA
  LocAddr1        As String * 64
  LocAddr2        As String * 64
  LocAddr3        As String * 64
  LocCity         As String * 64
  LocState        As String * 2
  LocZip          As String * 10
  LocCountyName   As String * 64
  LocCountry      As String * 3
  LocLatitude     As String * 15
  LocLongitude    As String * 15
'Tags
  TagLocBldType   As String * 64
  TagLocBldNumb   As String * 64
  TagLocSite      As String * 64
  TagLocFunding   As String * 64
  TagLocMainUse   As String * 64
  TagLocWatType   As String * 64
  TagLocArea      As String * 6
  TagLocIrrArea   As String * 6
  TagLocPopu      As String * 3
  TagLocWFR       As String * 1
  TagLocIrr       As String * 1
  TagLocYearBuilt As String * 4
  TagLocPool      As String * 1
  TagLocBathrooms As String * 3
  TagLocDistrict  As String * 64
  TagLocDHSCode   As String * 64
  TagLocParcNumb  As String * 32
  TagLocETJan     As String * 4
  TagLocETFeb     As String * 4
  TagLocETMar     As String * 4
  TagLocETApr     As String * 4
  TagLocETMay     As String * 4
  TagLocETJun     As String * 4
  TagLocETJul     As String * 4
  TagLocETAug     As String * 4
  TagLocETSep     As String * 4
  TagLocETOct     As String * 4
  TagLocETNov     As String * 4
  TagLocETDec     As String * 4

'Service Point
  SrvPntID        As String * 20
  SrvPntType      As String * 1
  SrvPntCycle     As String * 12
  SrvPntRoute     As String * 12
  SrvPntCsCode    As String * 24
  SrvPntCsCodeNm  As String * 24
  SrvPntLatitude  As String * 15
  SrvPntLongitude As String * 15
  SrvPntTimeZone  As String * 64
'Meter
  MtrID           As String * 40
  MtrSerNo        As String * 64
  MtrManufact     As String * 32
  MtrModel        As String * 64
  MtrSize         As String * 10
  MtrSizeUnit     As String * 6
  MtrNote         As String * 128
  MtrContFlow     As String * 1
  MtrRegNumb      As String * 1
  MtrRegUOM       As String * 12   'gallons, cubic feet, etc
  MtrRegReso      As String * 6
  MtrInstDate     As String * 10   'format is YYYY-MM-DD
  MtrInstStrRead  As String * 9
  MtrRemDate      As String * 10   'format is YYYY-MM-DD
  MtrRemRead      As String * 9
'Service Agreement
  SAStartDate     As String * 10
  SAEndDate       As String * 10
'Endpoint Config
  EndPointSerNo   As String * 20
  EndPointType    As String * 1
  EndPointInsDate As String * 10
  EndPointRemDate As String * 10
'Manual/Mobile Readings
  ManMobSequ      As String * 10
  ManMobAlertCode As String * 2
  ManMobHighRead  As String * 9
  ManMobLowRead   As String * 9
'Mics Info
  MIUse1          As String * 64
  MIUse2          As String * 64
End Type
  
Public Const BeaconHeader = "Account_ID|Account_First_Name|Account_Last_Name|Account_Full_Name|Account_Email|Account_Phone|" + _
                     "Account_Phone_Extension|Billing_Address_Line1|Billing_Address_Line2|Billing_Address_Line3|" + _
                     "Billing_City|Billing_State|Billing_Zip|Billing_Country|Person_ID|Account_Status|Account_Portal_Status|" + _
                     "Account_Billing_Cycle|Account_Paperless|Account_AutoPay|Account_BillerAutoPay|Location_ID|" + _
                     "Location_Name|Location_Address_Parity|Location_Address_Line1|Location_Address_Line2|" + _
                     "Location_Address_Line3|Location_City|Location_State|Location_Zip|Location_County_Name|" + _
                     "Location_Country|Location_Latitude|Location_Longitude|Location_Building_Type|" + _
                     "Location_Building_Number|Location_Site|Location_Funding|Location_Main_Use|Location_Water_Type|" + _
                     "Location_Area|Location_Irrigated_Area|Location_Population|Location_WFR|Location_Irrigation|" + _
                     "Location_Year_Built|Location_Pool|Location_Bathrooms|Location_District|Location_DHS_Code|" + _
                     "Location_Parcel_Number|Location_ET_Jan|Location_ET_Feb|Location_ET_Mar|Location_ET_Apr|" + _
                     "Location_ET_May|Location_ET_Jun|Location_ET_Jul|Location_ET_Aug|Location_ET_Sep|" + _
                     "Location_ET_Oct|Location_ET_Nov|Location_ET_Dec|Service_Point_ID|Service_Point_Type|" + _
                     "Service_Point_Cycle|Service_Point_Route|Service_Point_Class_Code|Service_Point_Class_Code_Normalized|" + _
                     "Service_Point_Latitude|Service_Point_Longitude|Service_Point_Timezone|Meter_ID|Meter_SN|" + _
                     "Meter_Manufacturer|Meter_Model|Meter_Size|Meter_Size_Unit|Meter_Note|Meter_Continuous_Flow|" + _
                     "Register_Number|Register_Unit_Of_Measure|Register_Resolution|Meter_Install_Date|" + _
                     "Meter_Install_Start_Read|Meter_Removal_Date|Meter_Removal_End_Read|SA_Start_Date|SA_End_Date|" + _
                     "Endpoint_SN|Endpoint_Type|Endpoint_Install_Date|Endpoint_Removal_Date|Read_Sequence|" + _
                     "Alert_Code|High_Read_Limit|Low_Read_Limit|Utility_Use_1|Utility_Use_2"



