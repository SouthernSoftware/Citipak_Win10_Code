Attribute VB_Name = "modPublicProcsAndVars"
Global glbIDNum As String
Global glbCaseNumber As String
Global glbName As String
Global glbArrestDate As String
Global glbAppPath As String

Public Function GetInmateCount(ByVal dtDate As Date, ByVal dtTime As Date, objDAODB As DAO.Database) As Long
    Dim lngCount As Long
    
    Dim strDate As String
    Dim strTime As String
    
    Dim rs As DAO.Recordset
    
    Set rs = objDAODB.OpenRecordset("select count(idnum) as InmateCount from [booking] where (dateofarrest <= #" & GetDateOnly(dtDate) & _
            "# AND timeofarrest <= #" & GetTimeOnly(dtTime) & "#) AND ((isnull(releasedate)) or (releasedate >= #" & _
            GetDateOnly(dtDate) & "# AND releasetime >= #" & GetTimeOnly(dtTime) & "# ))")
    
    If Not rs.EOF Then
        rs.MoveFirst
        If Not IsNull(rs("InmateCount")) Then
            lngCount = rs("inmatecount")
        End If
    End If
        
    GetInmateCount = lngCount
    objDAODB.Close
    
End Function

Public Function GetIDsFromName(ByVal blnEmployee As Boolean, ByVal blnInmate As Boolean, _
    ByRef lngMatchCountToBeReturned As Long, ByVal strName As String, objDAODB As DAO.Database) As DAO.Recordset
    
    Dim rs As DAO.Recordset
    
    If blnEmployee Then
        Set rs = objDAODB.OpenRecordset("select count(profidnum) as profidcount from [professionals] where profname = '" & strName & "'")
        If Not rs.EOF Then
            rs.MoveFirst
            lngMatchCountToBeReturned = rs("profidcount")
        End If
        Set rs = objDAODB.OpenRecordset("select PROFIDNUM from [professionals] where profname = '" & strName & "'")
    ElseIf blnInmate Then
        
        Set rs = objDAODB.OpenRecordset("select  count(idnum) as idCount from [booking] where sname = '" & strName & "'")
        If Not rs.EOF Then
            rs.MoveFirst
            If Not rs("IDCOUNT") = Null Then
                lngMatchCountToBeReturned = rs("idcount")
            End If
        End If
        
        Set rs = objDAODB.OpenRecordset("select IDNUM from [booking] where sname = '" & strName & "'")
    End If
    
    
    Set GetIDsFromName = rs
            
End Function
Public Function GetNameFromID(ByVal blnEmployee As Boolean, ByVal blnInmate As Boolean, _
    ByVal strID As String, objDAODB As DAO.Database) As DAO.Recordset
    
    Dim rs As Recordset
    
    
    Set rs = objDAODB.OpenRecordset("select profname from [professionals] where profidnum = '" & _
        strID & "'")
    
    If Not rs.EOF Then
        rs.MoveFirst
        rs.MoveNext
        If Not rs.EOF Then
            MsgBox "ID number " & strID & " returned Multiple Names in the database.  Duplicate Data Found.", vbOKOnly, "Genesis Error Log"
            Exit Function
        End If
    Else
        MsgBox "ID number " & strID & " not found in the database", vbOKOnly, "Genesis Error Log"
        Exit Function
    End If
    
    Set GetNameFromID = rs
    
    End Function

Public Sub SetListIndexToItemInList(objList As Control, ByVal strItemToMatch As String)

    Dim intx As Integer
    
    For intx = 0 To objList.ListCount
        If objList.List(intx) = strItemToMatch Then
            objList.ListIndex = intx
            Exit For
        End If
    Next intx
            
        
    
End Sub

Public Function TkBkSlshOffRtOfStr(ByVal strVal As String) As String
    strVal = Trim(strVal)
    If Trim(strVal) <> "" Then
        If Right$(strVal, 1) = "\" Then
            strVal = Mid$(strVal, 1, Len(strVal) - 1)
        End If
    End If
    
    TkBkSlshOffRtOfStr = strVal

End Function

Public Function GetDateOnly(dtDate As Date) As String
    Dim strVal As String
    
    strVal = CStr(dtDate)
    strVal = Format$(strVal, "mm/dd/yyyy")
    
    GetDateOnly = strVal
    

    
End Function

Public Function GetTimeOnly(dtTime As Date) As String

    Dim strVal As String
    
    strVal = CStr(dtTime)
    strVal = Format$(strVal, "Hh:Nn:Ss")
    
    GetTimeOnly = strVal
End Function

