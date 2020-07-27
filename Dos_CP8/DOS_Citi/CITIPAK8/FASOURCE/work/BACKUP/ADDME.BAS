
    DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)
    BCopy SSEG(OwnForm$(0, 0)), SADD(OwnForm$(0, 0)), VARSEG(UBOwnerRec(1)), VARPTR(UBOwnerRec(1)), UBOwnerRecLen, 0

