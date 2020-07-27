' These Sensus Layout Files are Spec'd Out Exactly to Gilbert SC

TYPE UBGilSensusReadRecType         ' File Layout for Sending Out Records
    CustLastName AS STRING * 25
    CustFirstName AS STRING * 25
    MeterID AS STRING * 8
    Account AS STRING * 8
    LowRead AS STRING * 8
    HighRead AS STRING * 8
    SensusType AS STRING * 1        ' B=Touch Read : M=Manual
    PastRead AS STRING * 8
    CurRead AS STRING * 8
    ServAddress AS STRING * 20
    LocationNumber AS STRING * 10
    Message AS STRING * 30
END TYPE

TYPE UBGilSensusGetReadRecType      ' File Layout For Retreiving Records
    CustLastName AS STRING * 25
    CustFirstName AS STRING * 25
    MeterID AS STRING * 8
    Account AS STRING * 8
    LowRead AS STRING * 8
    HighRead AS STRING * 8
    SensusType AS STRING * 1        ' B=Touch Read : M=Manual
    PastRead AS STRING * 8
    CurRead AS STRING * 8
    ServAddress AS STRING * 20
    LocationNumber AS STRING * 10
    Message AS STRING * 26
    DateRead AS STRING * 4
END TYPE


