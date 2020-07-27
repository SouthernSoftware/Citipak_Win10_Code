VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsControlResizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'----------------------------------------------------------------------------
' ClassName   : clsControlResizer
' Description : Automaticaly resize controls based on object tag
' Author      : CSN
' Date        : 07/13/2001
' Note        : 1. Put tag in controls <object>.tag rs=a|b|c|d|g
'               2. dim <variable> as new clsControlResizer
'               3. form_load,   call <variable>.Intialize(me [[width], [height] , [MinWidth] , [MinHeight]] )
'               4. form_resize, call <variable>.Resize(me)
'               5. form_unload, <variable> = nothing
'
'               Tags: a-left; b-top; c-right; d-buttom
'               In an MDI Child form, InitWidth and  initHeight are required
'
' Modification:
'
'   08/04/2001  - Supports resizing of columns in grid control. Just add rs=g
'                 grid must have columns collection and column(n).width property
'
'   08/06/2001  - Supports sstab control.
'----------------------------------------------------------------------------
Private Type ControlRecord
    a           As Integer
    b           As Integer
    c           As Integer
    d           As Integer
    Index       As Integer
    Level       As Integer
    GridWidth   As Integer
    GridTag     As String
End Type

Private arrControls()    As ControlRecord
Private intCount         As Integer
Private intMinWidth      As Integer
Private intMinHeight     As Integer
Private intCurrentWidth  As Integer
Private intCurrentHeight As Integer
Private Const sstabOffset = -75000


Public Sub Initialize(ByRef oForm As Object, Optional intInitWidth As Integer = -1, Optional intInitHeight As Integer = -1, Optional intWidth As Integer = 0, Optional intHeight As Integer = 0)
                    
    Dim i               As Integer
    Dim j               As Integer
    Dim strTag          As String
    Dim strLeft         As String
    Dim strTab          As String
    Dim crTemp          As ControlRecord
    Dim intTab          As Integer
    Dim intLeft         As Long
    
    On Error Resume Next
    intCount = 0
    intMinWidth = intWidth
    intMinHeight = intHeight
    If intInitWidth > -1 And intInitHeight > -1 Then
        oForm.Move 0, 0, intInitWidth, intInitHeight
    End If
    intCurrentWidth = 0
    intCurrentHeight = 0
    If oForm.MDIChild And intInitHeight = -1 And intInitWidth = -1 Then
        MsgBox "ControlResizer Error: You must supply initial width and height in an MDI child form!", vbCritical, "CSN"
        Exit Sub
    End If
    For i = 0 To oForm.Controls.Count - 1
        oForm.Controls(i).Tag = "rs=a|b|c|d|g"
        strTag = GetTag(oForm.Controls(i).Tag, "rs")
        If Left(strTag, 3) <> "" Then
            ReDim Preserve arrControls(0 To intCount)
            arrControls(intCount).a = -1
            arrControls(intCount).b = -1
            arrControls(intCount).c = -1
            arrControls(intCount).d = -1
            arrControls(intCount).Index = i
            arrControls(intCount).Level = levelcount(oForm.Controls(i))
            arrControls(intCount).GridWidth = -1
            arrControls(intCount).GridTag = ""
            
            intLeft = oForm.Controls(i).Left
            If intLeft < 0 Then
                intLeft = Abs(sstabOffset - intLeft)
            End If
            
            If InStr(1, strTag, "a") Then
                arrControls(intCount).a = intLeft
            End If
            If InStr(1, strTag, "b") Then
                arrControls(intCount).b = oForm.Controls(i).Top
            End If
            If InStr(1, strTag, "c") Then
                arrControls(intCount).c = oForm.Controls(i).Container.Width - intLeft - oForm.Controls(i).Width
            End If
            If InStr(1, strTag, "d") Then
                arrControls(intCount).d = oForm.Controls(i).Container.Height - oForm.Controls(i).Top - oForm.Controls(i).Height
            End If
            If InStr(1, strTag, "g") Then
                arrControls(intCount).GridWidth = oForm.Controls(i).Width
                For j = 0 To oForm.Controls(i).Columns.Count - 1
                    arrControls(intCount).GridTag = arrControls(intCount).GridTag & Format(oForm.Controls(i).Columns(j).Width, "00000")
                Next j
            End If
            intCount = intCount + 1
        End If
    Next i
    If intCount > 0 Then
        For i = 0 To intCount - 2
            For j = i + 1 To intCount - 1
                If arrControls(i).Level > arrControls(j).Level Then
                    crTemp = arrControls(i)
                    arrControls(i) = arrControls(j)
                    arrControls(j) = crTemp
                End If
            Next j
        Next i
    End If
End Sub

Private Function levelcount(oObject As Object) As Integer
    If TypeOf oObject.Container Is Form Then
        levelcount = 1
    Else
        levelcount = levelcount(oObject.Container) + 1
    End If
End Function

Public Sub Resize(ByRef oForm As Object)
    Dim i               As Integer
    Dim j               As Integer
    Dim intIndex        As Integer
    Dim intExcess       As Integer
    
    On Error Resume Next
    If oForm.WindowState = 0 Then
        If intMinWidth > 0 And oForm.Width < intMinWidth Then
            oForm.Width = intMinWidth
        End If
        If intMinHeight > 0 And oForm.Height < intMinHeight Then
            oForm.Height = intMinHeight
        End If
    End If
    If oForm.Height = intCurrentHeight And oForm.Width = intCurrentWidth Then
        Exit Sub
    End If
    For i = 0 To intCount - 1
        intIndex = arrControls(i).Index
        If arrControls(i).a > -1 Then
            If oForm.Controls(intIndex).Left < 0 Then
                oForm.Controls(intIndex).Left = sstabOffset + arrControls(i).a
            Else
                oForm.Controls(intIndex).Left = arrControls(i).a
            End If
        End If
        If arrControls(i).b > -1 Then
            oForm.Controls(intIndex).Top = arrControls(i).b
        End If
        If arrControls(i).c > -1 Then
            If arrControls(i).a = -1 Then
                If oForm.Controls(intIndex).Left < 0 Then
                    oForm.Controls(intIndex).Left = sstabOffset + (oForm.Controls(intIndex).Container.Width - oForm.Controls(intIndex).Width - arrControls(i).c)
                Else
                    oForm.Controls(intIndex).Left = oForm.Controls(intIndex).Container.Width - oForm.Controls(intIndex).Width - arrControls(i).c
                End If
            Else
                If oForm.Controls(intIndex).Left < 0 Then
                    oForm.Controls(intIndex).Width = sstabOffset + (oForm.Controls(intIndex).Container.Width - oForm.Controls(intIndex).Left - arrControls(i).c)
                Else
                    oForm.Controls(intIndex).Width = oForm.Controls(intIndex).Container.Width - oForm.Controls(intIndex).Left - arrControls(i).c
                End If
            End If
        End If
        If arrControls(i).d > -1 Then
            If arrControls(i).b = -1 Then
                oForm.Controls(intIndex).Top = CheckNegative(oForm.Controls(intIndex).Container.Height - oForm.Controls(intIndex).Height - arrControls(i).d)
            Else
                oForm.Controls(intIndex).Height = CheckNegative(oForm.Controls(intIndex).Container.Height - oForm.Controls(intIndex).Top - arrControls(i).d)
            End If
        End If
        If arrControls(i).GridWidth > -1 Then
            If oForm.Controls(intIndex).Width > arrControls(i).GridWidth Then
                intExcess = (oForm.Controls(intIndex).Width - arrControls(i).GridWidth) / oForm.Controls(intIndex).Columns.Count
            Else
                intExcess = 0
            End If
            For j = 0 To oForm.Controls(intIndex).Columns.Count - 1
                oForm.Controls(intIndex).Columns(j).Width = Val(Mid(arrControls(i).GridTag, j * 5 + 1, 5)) + intExcess
            Next j
        End If
    Next i
End Sub

Private Function CheckNegative(ByRef intExpr As Integer) As Integer
    If intExpr < 0 Then
        CheckNegative = 0
    Else
        CheckNegative = intExpr
    End If
End Function

Private Function GetTag(ByVal strTag As String, ByVal strKey As String) As String
    Dim nPos1 As Integer
    Dim nPos2 As Integer
    GetTag = ""
    strTag = strTag & ";"
    nPos1 = InStr(1, strTag, strKey & "=")
    If nPos1 > 0 Then
        GetTag = Mid(strTag, nPos1, InStr(nPos1, strTag, ";") - nPos1)
    End If
End Function



