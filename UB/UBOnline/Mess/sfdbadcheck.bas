Public Sub Savebadcheck(frm as form, pathvar as string)
dim db as database, rs as recordset
set db = opendatabase(+"incident.mdb")
set rs = db.openrecordset("select * from badcheck where incidentnumber = '"+frm.incidentnumber+"'")
if rs.eof then
  rs.addnew
else
  rs.edit
end if
rs("incidentnumber") = frm.incidentnumber
rs("cname") = frm.cname
rs("caddress") = frm.caddress
rs("cphone") = frm.cphone
rs("incidentdate") = frm.incidentdate
rs("incidenttime") = frm.incidenttime
rs("dateofoffense") = frm.dateofoffense
rs("incidentlocation") = frm.incidentlocation
rs("highway") = frm.highway
rs("commercial") = frm.commercial
rs("scvstation") = frm.scvstation
rs("chainstore") = frm.chainstore
rs("residence") = frm.residence
rs("bank") = frm.bank
rs("other") = frm.other
rs("otherspecify") = frm.otherspecify
rs("vname") = frm.vname
rs("vaddress") = frm.vaddress
rs("vrace") = frm.vrace
rs("vsex") = frm.vsex
rs("vage") = frm.vage
rs("vhphone") = frm.vhphone
rs("vwphone") = frm.vwphone
rs("suspect") = frm.suspect
rs("wanted") = frm.wanted
rs("warrant") = frm.warrant
rs("arrest") = frm.arrest
rs("sname") = frm.sname
rs("saddress") = frm.saddress
rs("srace") = frm.srace
rs("ssex") = frm.ssex
rs("sage") = frm.sage
rs("drivers") = frm.drivers
rs("ssn") = frm.ssn
rs("sdateofbirth") = frm.sdateofbirth
rs("sheight") = frm.sheight
rs("sweight") = frm.sweight
rs("shair") = frm.shair
rs("seyes") = frm.seyes
rs("totalarrested") = frm.totalarrested
rs("nearyes") = frm.nearyes
rs("nearno") = frm.nearno
rs("checknumbers") = frm.checknumbers
rs("bankname") = frm.bankname
rs("status") = frm.status
rs("comments") = frm.comments
rs("theft") = frm.theft
rs("recovery") = frm.recovery
rs("active") = frm.active
rs("cleared") = frm.cleared
rs("reportingofficer") = frm.reportingofficer
rs("reportingofficerunit") = frm.reportingofficerunit
rs("approvingofficer") = frm.approvingofficer
rs("approvingofficerunit") = frm.approvingofficerunit
rs("jurisdiction") = frm.jurisdiction
rs("idnumber") = frm.idnumber
rs("driversstate") = frm.driversstate
rs("ccity") = frm.ccity
rs("cstate") = frm.cstate
rs("czipcode") = frm.czipcode
rs("vcity") = frm.vcity
rs("vstate") = frm.vstate
rs("vzipcode") = frm.vzipcode
rs("scity") = frm.scity
rs("sstate") = frm.sstate
rs("szipcode") = frm.szipcode
rs("ORINUMBER") = frmlogin.orinumber
rs("userid") = frmlogin.userid
rs("userfullname") = frmlogin.userfullname
rs("udate") = date$
rs("utime") = time$
rs("checkamount") = frm.checkamount
rs.update
rs.close
db.close
End Sub
Public Sub Findbadcheck(frm as form, pathvar as string)
dim db as database, rs as recordset
set db = opendatabase(+"incident.mdb")
set rs = db.openrecordset("select * from badcheck where incidentnumber = '"+frm.incidentnumber+"'")
call clearbadcheck(frm)
if rs.eof then
  exit sub
end if
if not isnull(rs("incidentnumber")) then
 frm.incidentnumber = rs("incidentnumber")
end if
if not isnull(rs("cname")) then
 frm.cname = rs("cname")
end if
if not isnull(rs("caddress")) then
 frm.caddress = rs("caddress")
end if
if not isnull(rs("cphone")) then
 frm.cphone = rs("cphone")
end if
if not isnull(rs("incidentdate")) then
 frm.incidentdate = rs("incidentdate")
end if
if not isnull(rs("incidenttime")) then
 frm.incidenttime = rs("incidenttime")
end if
if not isnull(rs("dateofoffense")) then
 frm.dateofoffense = rs("dateofoffense")
end if
if not isnull(rs("incidentlocation")) then
 frm.incidentlocation = rs("incidentlocation")
end if
if not isnull(rs("highway")) then
 frm.highway = rs("highway")
end if
if not isnull(rs("commercial")) then
 frm.commercial = rs("commercial")
end if
if not isnull(rs("scvstation")) then
 frm.scvstation = rs("scvstation")
end if
if not isnull(rs("chainstore")) then
 frm.chainstore = rs("chainstore")
end if
if not isnull(rs("residence")) then
 frm.residence = rs("residence")
end if
if not isnull(rs("bank")) then
 frm.bank = rs("bank")
end if
if not isnull(rs("other")) then
 frm.other = rs("other")
end if
if not isnull(rs("otherspecify")) then
 frm.otherspecify = rs("otherspecify")
end if
if not isnull(rs("vname")) then
 frm.vname = rs("vname")
end if
if not isnull(rs("vaddress")) then
 frm.vaddress = rs("vaddress")
end if
if not isnull(rs("vrace")) then
 frm.vrace = rs("vrace")
end if
if not isnull(rs("vsex")) then
 frm.vsex = rs("vsex")
end if
if not isnull(rs("vage")) then
 frm.vage = rs("vage")
end if
if not isnull(rs("vhphone")) then
 frm.vhphone = rs("vhphone")
end if
if not isnull(rs("vwphone")) then
 frm.vwphone = rs("vwphone")
end if
if not isnull(rs("suspect")) then
 frm.suspect = rs("suspect")
end if
if not isnull(rs("wanted")) then
 frm.wanted = rs("wanted")
end if
if not isnull(rs("warrant")) then
 frm.warrant = rs("warrant")
end if
if not isnull(rs("arrest")) then
 frm.arrest = rs("arrest")
end if
if not isnull(rs("sname")) then
 frm.sname = rs("sname")
end if
if not isnull(rs("saddress")) then
 frm.saddress = rs("saddress")
end if
if not isnull(rs("srace")) then
 frm.srace = rs("srace")
end if
if not isnull(rs("ssex")) then
 frm.ssex = rs("ssex")
end if
if not isnull(rs("sage")) then
 frm.sage = rs("sage")
end if
if not isnull(rs("drivers")) then
 frm.drivers = rs("drivers")
end if
if not isnull(rs("ssn")) then
 frm.ssn = rs("ssn")
end if
if not isnull(rs("sdateofbirth")) then
 frm.sdateofbirth = rs("sdateofbirth")
end if
if not isnull(rs("sheight")) then
 frm.sheight = rs("sheight")
end if
if not isnull(rs("sweight")) then
 frm.sweight = rs("sweight")
end if
if not isnull(rs("shair")) then
 frm.shair = rs("shair")
end if
if not isnull(rs("seyes")) then
 frm.seyes = rs("seyes")
end if
if not isnull(rs("totalarrested")) then
 frm.totalarrested = rs("totalarrested")
end if
if not isnull(rs("nearyes")) then
 frm.nearyes = rs("nearyes")
end if
if not isnull(rs("nearno")) then
 frm.nearno = rs("nearno")
end if
if not isnull(rs("checknumbers")) then
 frm.checknumbers = rs("checknumbers")
end if
if not isnull(rs("bankname")) then
 frm.bankname = rs("bankname")
end if
if not isnull(rs("status")) then
 frm.status = rs("status")
end if
if not isnull(rs("comments")) then
 frm.comments = rs("comments")
end if
if not isnull(rs("theft")) then
 frm.theft = rs("theft")
end if
if not isnull(rs("recovery")) then
 frm.recovery = rs("recovery")
end if
if not isnull(rs("active")) then
 frm.active = rs("active")
end if
if not isnull(rs("cleared")) then
 frm.cleared = rs("cleared")
end if
if not isnull(rs("reportingofficer")) then
 frm.reportingofficer = rs("reportingofficer")
end if
if not isnull(rs("reportingofficerunit")) then
 frm.reportingofficerunit = rs("reportingofficerunit")
end if
if not isnull(rs("approvingofficer")) then
 frm.approvingofficer = rs("approvingofficer")
end if
if not isnull(rs("approvingofficerunit")) then
 frm.approvingofficerunit = rs("approvingofficerunit")
end if
if not isnull(rs("jurisdiction")) then
 frm.jurisdiction = rs("jurisdiction")
end if
if not isnull(rs("idnumber")) then
 frm.idnumber = rs("idnumber")
end if
if not isnull(rs("driversstate")) then
 frm.driversstate = rs("driversstate")
end if
if not isnull(rs("ccity")) then
 frm.ccity = rs("ccity")
end if
if not isnull(rs("cstate")) then
 frm.cstate = rs("cstate")
end if
if not isnull(rs("czipcode")) then
 frm.czipcode = rs("czipcode")
end if
if not isnull(rs("vcity")) then
 frm.vcity = rs("vcity")
end if
if not isnull(rs("vstate")) then
 frm.vstate = rs("vstate")
end if
if not isnull(rs("vzipcode")) then
 frm.vzipcode = rs("vzipcode")
end if
if not isnull(rs("scity")) then
 frm.scity = rs("scity")
end if
if not isnull(rs("sstate")) then
 frm.sstate = rs("sstate")
end if
if not isnull(rs("szipcode")) then
 frm.szipcode = rs("szipcode")
end if
if not isnull(rs("checkamount")) then
 frm.checkamount = rs("checkamount")
end if
rs.close
db.close
End Sub
Public Sub Deletebadcheck(frm as form, pathvar as string)
msg = msgbox("Are you sure you want to delete this record?",4,"Genesis Information Log")
if msg <> 6 then
  exit sub
end if
dim db as database, rs as recordset
set db = opendatabase(+"incident.mdb")
set rs = db.openrecordset("select * from badcheck where incidentnumber = '"+frm.incidentnumber+"'")
while not rs.eof
  rs.delete
  rs.movenext
wend
rs.close
db.close
End Sub
Public Sub Clearbadcheck(frm as form)
frm.incidentnumber = ""
frm.cname = ""
frm.caddress = ""
frm.cphone = ""
frm.incidentdate = ""
frm.incidenttime = ""
frm.dateofoffense = ""
frm.incidentlocation = ""
frm.highway = 0
frm.commercial = 0
frm.scvstation = 0
frm.chainstore = 0
frm.residence = 0
frm.bank = 0
frm.other = 0
frm.otherspecify = ""
frm.vname = ""
frm.vaddress = ""
frm.vrace = ""
frm.vsex = ""
frm.vage = ""
frm.vhphone = ""
frm.vwphone = ""
frm.suspect = 0
frm.wanted = 0
frm.warrant = 0
frm.arrest = 0
frm.sname = ""
frm.saddress = ""
frm.srace = ""
frm.ssex = ""
frm.sage = ""
frm.drivers = ""
frm.ssn = ""
frm.sdateofbirth = ""
frm.sheight = ""
frm.sweight = ""
frm.shair = ""
frm.seyes = ""
frm.totalarrested = 0
frm.nearyes = False
frm.nearno = False
frm.checknumbers = ""
frm.bankname = ""
frm.status = ""
frm.comments = 0
frm.theft = 0
frm.recovery = 0
frm.active = 0
frm.cleared = 0
frm.reportingofficer = ""
frm.reportingofficerunit = ""
frm.approvingofficer = ""
frm.approvingofficerunit = ""
frm.jurisdiction = ""
frm.idnumber = ""
frm.driversstate = ""
frm.ccity = ""
frm.cstate = ""
frm.czipcode = ""
frm.vcity = ""
frm.vstate = ""
frm.vzipcode = ""
frm.scity = ""
frm.sstate = ""
frm.szipcode = ""
frm.checkamount = ""
End Sub
