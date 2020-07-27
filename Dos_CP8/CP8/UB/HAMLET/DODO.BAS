              AvgUse# = UBCustRec(1).LocMeters(MtrCnt&).AvgUse
              UseCnt = UBCustRec(1).LocMeters(MtrCnt&).UseCnt
              IF AvgUse# < 0 THEN
                AvgUse# = 1
              END IF
              IF UseCnt < 0 THEN
                UseCnt = 1
              END IF
              MeterConsp& = UBCustRec(1).LocMeters(MtrCnt&).CurRead - UBCustRec(1).LocMeters(MtrCnt&).PrevRead
              IF MeterConsp& < 0 THEN
                MaxMeterAmt& = 10& ^ (LEN(STR$(UBCustRec(1).LocMeters(MtrCnt&).PrevRead)) - 1)
                MeterConsp& = (MaxMeterAmt& - UBCustRec(1).LocMeters(MtrCnt&).PrevRead) + UBCustRec(1).LocMeters(MtrCnt&).CurRead
              END IF

              IF UseHigh THEN
                HiAvg# = (AvgUse# * (HighVar# / 100))
                IF MeterConsp& > HiAvg# THEN
                  HighFlag = True
                END IF
              END IF

