TYPE ItemRecType
     ITEMNUMBER AS LONG
     DESC1 AS STRING * 45
     DESC2 AS STRING * 45
     Vendor AS STRING * 30
     VItem AS STRING * 20
     DeptNumb AS STRING * 8
     Cost AS DOUBLE
     Unit AS STRING * 10
     QOH AS SINGLE
     QonOrder AS DOUBLE
     UsedYTD AS DOUBLE
     ReOrder AS DOUBLE
     Note1 AS STRING * 45
     Note2 AS STRING * 45
     Deleted AS STRING * 1              'Y if Deleted
     Future AS STRING * 383
END TYPE

TYPE ICTransRecType
   TransDate AS INTEGER
   TransAmount AS DOUBLE
   TransType AS INTEGER         '1=Up Adj 2=Dwn Adj 3=Redorder 4=Used 5=Purchased
   TransDesc1 AS STRING * 40
   TransDesc2 AS STRING * 40
   TransItemRecord AS LONG
   TransNext AS LONG
   TransPad AS STRING * 28
END TYPE

