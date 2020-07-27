Attribute VB_Name = "modZoom"
Option Explicit

Global zoomindex As Integer
Sub GetZoom(zoomlabel As Integer)
'Set up the print previews zoom

        Select Case zoomlabel
            Case 0
                spreadpreview.fpSpreadPreview1.PageViewType = 2
                spreadpreview.fpSpreadPreview1.PageViewPercentage = 200
            
            Case 1
                spreadpreview.fpSpreadPreview1.PageViewType = 2
                spreadpreview.fpSpreadPreview1.PageViewPercentage = 150

            Case 2
                spreadpreview.fpSpreadPreview1.PageViewType = 2
                spreadpreview.fpSpreadPreview1.PageViewPercentage = 100

            Case 3
                spreadpreview.fpSpreadPreview1.PageViewType = 2
                spreadpreview.fpSpreadPreview1.PageViewPercentage = 75

            Case 4
                spreadpreview.fpSpreadPreview1.PageViewType = 2
                spreadpreview.fpSpreadPreview1.PageViewPercentage = 50

            Case 5
                spreadpreview.fpSpreadPreview1.PageViewType = 2
                spreadpreview.fpSpreadPreview1.PageViewPercentage = 25

            Case 6
                spreadpreview.fpSpreadPreview1.PageViewType = 2
                spreadpreview.fpSpreadPreview1.PageViewPercentage = 10

            Case 7
                spreadpreview.fpSpreadPreview1.PageViewType = 3
                
            Case 8
                spreadpreview.fpSpreadPreview1.PageViewType = 4
                
            Case 9
                spreadpreview.fpSpreadPreview1.PageViewType = 0
                
            Case 10
                spreadpreview.fpSpreadPreview1.PageViewType = 5
                spreadpreview.fpSpreadPreview1.PageMultiCntH = 2
                spreadpreview.fpSpreadPreview1.PageMultiCntV = 1
                
            Case 11
                spreadpreview.fpSpreadPreview1.PageViewType = 5
                spreadpreview.fpSpreadPreview1.PageMultiCntH = 3
                spreadpreview.fpSpreadPreview1.PageMultiCntV = 1
                
            Case 12
                spreadpreview.fpSpreadPreview1.PageViewType = 5
                spreadpreview.fpSpreadPreview1.PageMultiCntH = 2
                spreadpreview.fpSpreadPreview1.PageMultiCntV = 2
                
            Case 13
                spreadpreview.fpSpreadPreview1.PageViewType = 5
                spreadpreview.fpSpreadPreview1.PageMultiCntH = 3
                spreadpreview.fpSpreadPreview1.PageMultiCntV = 2

        End Select
      
End Sub


