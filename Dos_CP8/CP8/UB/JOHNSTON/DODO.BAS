              MtrNumb$ = QPTrim$(UBCustRec(1).LocMeters(mChk).MTRNUM)
              IF LEN(MtrNumb$) = 0 THEN
                MtrNumb$ = "???"
              END IF

