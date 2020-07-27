      GET #TransFile, TaxIntRec(1).billrec, TaxTrans(1) 'get bill trans
      TaxTrans(1).Revenue.interest = Round#(TaxTrans(1).Revenue.interest + TaxIntRec(1).Amount)
      PUT #TransFile, TaxIntRec(1).billrec, TaxTrans(1) 'put it back
      'Now make a new clean transaction
      REDIM TaxTrans(1) AS TaxTransactionType
      TaxTrans(1).TransDate = Date2Num%(DATE$)
      TaxTrans(1).TAXYEAR = TaxIntRec(1).TAXYEAR
      TaxTrans(1).TranType = 4       '4=Interest
      TaxTrans(1).BillType = "I"     'R=Real P=Personal Property C=Combined (NC/GA)
      TaxTrans(1).Amount = TaxIntRec(1).Amount  'Total Transaction Amount
      TaxTrans(1).Revenue.interest = TaxIntRec(1).Amount
      TaxTrans(1).Description = "Tax Int on Bill# " + QPTrim$(TaxIntRec(1).BillNumber)
      TaxTrans(1).Posted2GL = "N"
      TaxTrans(1).CustomerRec = TaxIntRec(1).CustRec
      TaxTrans(1).LastTrans = 0
      TaxTrans(1).BelongTo = TaxIntRec(1).billrec
      LSET TaxTrans(1).Padding = ""
      'Increment Transaction File Record Count
      NextRecord& = (LOF(TransFile) / LEN(TaxTrans(1))) + 1
      PUT #TransFile, NextRecord&, TaxTrans(1)
      'Update the Customer Pointers Now
      GET #CustFile, TaxIntRec(1).CustRec, TaxCust(1)

