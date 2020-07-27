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

