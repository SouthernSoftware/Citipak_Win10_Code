Attribute VB_Name = "modSpellGmr"
Public blnFrmSpellShowing As Boolean
Private strSpellDoc As String ' the string to checkspelling
Private objNarrative As Object
Public Sub BeginSpellCheck(strVal As String, objControl As Object)
Dim DONE As Boolean
Set objNarrative = objControl
Call CheckSpelling(DONE, strVal)
If DONE Then
    msg = MsgBox("Spelling check complete.", 48, "Genesis Information Log")
End If
End Sub

Private Sub CheckSpelling(blnDone As Boolean, strToCheck As String)
Dim wd As New Word.Application
Dim wdsp As Word.SpellingSuggestions

On Error GoTo cmdCheckErr
strSpellDoc = strToCheck
GETOUT% = 0
Unload frmSpell
wd.WindowState = wdWindowStateMinimize
wd.Options.CheckGrammarWithSpelling = True
While strSpellDoc > "" And GETOUT% = 0
    'wd.Visible = False
    While Left$(strSpellDoc, 1) = " "
        strSpellDoc = Mid$(strSpellDoc, 2)
    Wend
    If strSpellDoc = "" Then
        GoTo gowend
    End If
    stopper% = Len(strSpellDoc) + 1
    For t% = 1 To Len(strSpellDoc)
        If InStr("!()[{]};:,./? " + Chr$(34), Mid$(strSpellDoc, t%, 1)) Then
            stopper% = t%
            t% = Len(strSpellDoc)
        End If
    Next t%
    frmSpell.checkword = LCase(Left$(strSpellDoc, stopper% - 1))
    
    
    If stopper% < Len(strSpellDoc) Then
        strSpellDoc = Mid$(strSpellDoc, stopper% + 1)
    Else
        strSpellDoc = ""
    End If
    
    

    
    If frmSpell.checkword > "" Then
        wd.Documents.add
        Set wdsp = wd.GetSpellingSuggestions(frmSpell.checkword)
       
        If wdsp.Count > 0 Or wdsp.SpellingErrorType = wdSpellingNotInDictionary Then
            frmSpell.lstsuggestions.clear
            frmSpell.Show
            'RLB code
            'lstframe.Top = objnarrative.Top - lstframe.Height
            'lstframe.Left = objnarrative.Left + CLng(lstframe.Width * 0.5)
            '***********
            GETOUT% = 1
        Else
            Unload frmSpell
        End If
        
        For i% = 1 To wdsp.Count
            frmSpell.lstsuggestions.AddItem wdsp(i%).Name
        Next i%
        If wdsp.SpellingErrorType = wdSpellingNotInDictionary And wdsp.Count = 0 Then
            frmSpell.lstsuggestions.AddItem "Not found in dictionary."
        End If
        wd.Documents.Close
        'wd.Visible = False
    Else
        
    End If
gowend:
Wend
wd.Quit
Set wd = Nothing
Screen.MousePointer = 0
If strSpellDoc = "" And (Not blnFrmSpellShowing) Then
    blnDone = True
Else
    blnDone = False
End If
Exit Sub
cmdCheckErr:
'MsgBox Err.description
Resume Next
End Sub
Public Sub SpellChange() ' used by frmspell
Dim DONE As Boolean
If frmSpell.lstsuggestions.ListIndex = -1 Then
    Exit Sub
End If
For t% = 1 To Len(objNarrative.Text)
    If UCase(Mid$(objNarrative.Text, t%, Len(frmSpell.checkword))) = UCase(frmSpell.checkword) Then
        If UCase(Mid$(objNarrative.Text, t%, Len(frmSpell.checkword))) = Mid$(objNarrative.Text, t%, Len(frmSpell.checkword)) Then
            frmSpell.lstsuggestions.List(frmSpell.lstsuggestions.ListIndex) = UCase(frmSpell.lstsuggestions.List(frmSpell.lstsuggestions.ListIndex))
        End If
        objNarrative.Text = Left$(objNarrative.Text, t% - 1) + frmSpell.lstsuggestions.List(frmSpell.lstsuggestions.ListIndex) + Mid$(objNarrative.Text, t% + Len(frmSpell.checkword))
        t% = t% + Len(frmSpell.checkword)
    End If
Next t%

'RLB CODE
For t% = 1 To Len(strSpellDoc)
    If UCase(Mid$(strSpellDoc, t%, Len(frmSpell.checkword))) = UCase(frmSpell.checkword) Then
        If UCase(Mid$(strSpellDoc, t%, Len(frmSpell.checkword))) = Mid$(strSpellDoc, t%, Len(frmSpell.checkword)) Then
            frmSpell.lstsuggestions.List(frmSpell.lstsuggestions.ListIndex) = UCase(frmSpell.lstsuggestions.List(frmSpell.lstsuggestions.ListIndex))
        End If
        strSpellDoc = Left$(strSpellDoc, t% - 1) + frmSpell.lstsuggestions.List(frmSpell.lstsuggestions.ListIndex) + Mid$(strSpellDoc, t% + Len(frmSpell.checkword))
        t% = t% + Len(frmSpell.checkword)
    End If
Next t%
'*********


Call CheckSpelling(DONE, strSpellDoc)
If DONE Then
    msg = MsgBox("Spelling check complete.", 48, "Genesis Information Log")
End If
End Sub
Public Sub SkipWord() 'used by frmspell
Dim DONE As Boolean

Call CheckSpelling(DONE, strSpellDoc)
If DONE Then
    msg = MsgBox("Spelling check complete.", 48, "Genesis Information Log")
End If
End Sub
