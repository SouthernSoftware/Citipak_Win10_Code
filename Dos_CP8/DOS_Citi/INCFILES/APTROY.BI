'OPEN "apvendor.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #1 LEN = 204
type Vendor6RecType
 VenName as string * 30  '30 AS venname$,
 address as string * 30  '30 AS address$,
 city as string * 22     '22 AS CitY$,
 St as string * 2        '2 AS St$,
 Zip as string * 10      '10 AS ZiP$,
 PayTo as string * 30    '30 AS payto$,
 payadd as string * 30   '30 AS payadd$,
 paycity as string * 22  '22 AS paycity$,
 payst  as string * 2    '2 AS payst$,
 payzip as string * 10   '10 AS payzip$,
 CURBAL as Double        '8 AS CURBAL$,
 first  as single        '4 AS first$,
 last   as single        '4 AS last$
end type
