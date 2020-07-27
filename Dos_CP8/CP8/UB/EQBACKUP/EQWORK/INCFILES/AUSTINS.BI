TYPE Customer
        Acct AS STRING * 12
        Name AS STRING * 25
        Add1 AS STRING * 25
        Add2 AS STRING * 25
        City AS STRING * 15
        State AS STRING * 2
        Zip AS STRING * 10
        Contact AS STRING * 25
        LastPurchase AS STRING * 6
        Pad AS STRING * 112
END TYPE

TYPE CustomerImport
        Nul AS STRING * 3
        Name AS STRING * 25
        Acct AS STRING * 12
        Nul2 AS STRING * 1
        Add1 AS STRING * 25
        Add2 AS STRING * 25
        City AS STRING * 15
        State AS STRING * 2
        Zip AS STRING * 10
        Nul1 AS STRING * 12
        Contact AS STRING * 25
        Nul5 AS STRING * 499
        LastDate AS STRING * 6
END TYPE

