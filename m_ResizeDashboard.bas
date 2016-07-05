Attribute VB_Name = "m_ResizeDashboard"
'---------------------------------------------------------------------------------------
' Module    : m_ResizeDashboard
' Author    : Jon Acampora, ExcelCampus.com
' Date      : 7/20/2015
' Purpose   : Move and resize shapes in a dashboard to fit on different screen sizes
'---------------------------------------------------------------------------------------

Option Explicit

Sub Snapshot_View()
'Lists the location and size of all the shapes
'on the activesheet in the immediate window.
'Copy the code into one of the View macros.
'CLEAR THE IMMEDIATE WINDOW BEFORE RUNNING THIS

Dim shp As Shape

    Debug.Print "'Window Width: " & ActiveWindow.Width
    Debug.Print "'Window Height: " & ActiveWindow.Height
    Debug.Print ""
    Debug.Print "    With Sheet1"
    For Each shp In ActiveSheet.Shapes
        Debug.Print "        .Shapes(" & Chr(34) & shp.Name & Chr(34) & ").Top = " & shp.Top
        Debug.Print "        .Shapes(" & Chr(34) & shp.Name & Chr(34) & ").Left = " & shp.Left
        Debug.Print "        .Shapes(" & Chr(34) & shp.Name & Chr(34) & ").Height = " & shp.Height
        Debug.Print "        .Shapes(" & Chr(34) & shp.Name & Chr(34) & ").Width = " & shp.Width
    Next shp
    Debug.Print "    End With"
    
End Sub

Sub Change_View()
'Call the view macro based on the height or width of the window.
'Call by the Workbook_Open and Workbook_WindowResize events
'in the ThisWorkbook module.

Dim dWidth As Double
Dim dHeight As Double

    dWidth = ActiveWindow.Width
    dHeight = ActiveWindow.Height
    
    'Uncomment the next line to change the zoom to 100% before resizing
    'ActiveWindow.Zoom = 100
    
    'Start with largest screensize
    If dHeight >= 535 And dWidth >= 955 Then
        Call View_Laptop_H720px
        
    ElseIf dHeight >= 498 And dWidth >= 855.75 Then
        Call View_Small_Laptop
        
    ElseIf dWidth >= 650 Then
        'Smaller screens can have vertical scroll
        'Don't consider the height of the screen
        Call View_Tablet
        
    ElseIf dWidth >= 475 Then
        Call View_Narrow
        
    End If

End Sub

Sub View_Laptop_H720px()
'Window Width: 960
'Window Height: 540

    With Sheet1
        .Shapes("Chart 1").Top = 32.25
        .Shapes("Chart 1").Left = 14.24976
        .Shapes("Chart 1").Height = 170.25
        .Shapes("Chart 1").Width = 529.5002
        .Shapes("Chart 3").Top = 211.1093
        .Shapes("Chart 3").Left = 553.5002
        .Shapes("Chart 3").Height = 178.8907
        .Shapes("Chart 3").Width = 228.75
        .Shapes("Chart 7").Top = 213
        .Shapes("Chart 7").Left = 14.99992
        .Shapes("Chart 7").Height = 172.5
        .Shapes("Chart 7").Width = 289.5001
        .Shapes("Chart 8").Top = 33
        .Shapes("Chart 8").Left = 553.5
        .Shapes("Chart 8").Height = 170.25
        .Shapes("Chart 8").Width = 228
        .Shapes("Chart 9").Top = 211.9826
        .Shapes("Chart 9").Left = 315.7501
        .Shapes("Chart 9").Height = 175.7674
        .Shapes("Chart 9").Width = 228
        .Shapes("Salesperson").Top = 156.7501
        .Shapes("Salesperson").Left = 798.0002
        .Shapes("Salesperson").Height = 193.5
        .Shapes("Salesperson").Width = 101.2499
        .Shapes("Region 2").Top = 33.75008
        .Shapes("Region 2").Left = 797.2502
        .Shapes("Region 2").Height = 116.9999
        .Shapes("Region 2").Width = 101.2499
        .Shapes("TextBox 12").Top = 6
        .Shapes("TextBox 12").Left = 10.5
        .Shapes("TextBox 12").Height = 27
        .Shapes("TextBox 12").Width = 237
        .Shapes("TextBox 13").Top = 6
        .Shapes("TextBox 13").Left = 681.75
        .Shapes("TextBox 13").Height = 24
        .Shapes("TextBox 13").Width = 119.8235
        .Shapes("Picture 17").Top = 11.99992
        .Shapes("Picture 17").Left = 670.4998
        .Shapes("Picture 17").Height = 17.25
        .Shapes("Picture 17").Width = 17.25
    End With

End Sub

Sub View_Small_Laptop()
'Window Width: 855.75
'Window Height: 498

    With Sheet1
        .Shapes("Chart 1").Top = 32.25008
        .Shapes("Chart 1").Left = 14.24976
        .Shapes("Chart 1").Height = 154.5
        .Shapes("Chart 1").Width = 447.7502
        .Shapes("Chart 3").Top = 195.3593
        .Shapes("Chart 3").Left = 516.0001
        .Shapes("Chart 3").Height = 153.3907
        .Shapes("Chart 3").Width = 181.5
        .Shapes("Chart 7").Top = 195.3593
        .Shapes("Chart 7").Left = 14.99992
        .Shapes("Chart 7").Height = 153.3907
        .Shapes("Chart 7").Width = 273.7501
        .Shapes("Chart 8").Top = 32.25008
        .Shapes("Chart 8").Left = 470.25
        .Shapes("Chart 8").Height = 154.5
        .Shapes("Chart 8").Width = 228
        .Shapes("Chart 9").Top = 195.3593
        .Shapes("Chart 9").Left = 293.25
        .Shapes("Chart 9").Height = 153.3907
        .Shapes("Chart 9").Width = 217.5
        .Shapes("Salesperson").Top = 156.0001
        .Shapes("Salesperson").Left = 706.5002
        .Shapes("Salesperson").Height = 193.5
        .Shapes("Salesperson").Width = 101.2499
        .Shapes("Region 2").Top = 33.00008
        .Shapes("Region 2").Left = 705.7502
        .Shapes("Region 2").Height = 116.9999
        .Shapes("Region 2").Width = 101.2499
        .Shapes("TextBox 12").Top = 6
        .Shapes("TextBox 12").Left = 10.5
        .Shapes("TextBox 12").Height = 27
        .Shapes("TextBox 12").Width = 237
        .Shapes("TextBox 13").Top = 9
        .Shapes("TextBox 13").Left = 601.5
        .Shapes("TextBox 13").Height = 24
        .Shapes("TextBox 13").Width = 119.8235
        .Shapes("Picture 17").Top = 14.99992
        .Shapes("Picture 17").Left = 590.2498
        .Shapes("Picture 17").Height = 17.25
        .Shapes("Picture 17").Width = 17.25
    End With


End Sub

Sub View_Tablet()
'Window Width: 636.75
'Window Height: 472.5

    With Sheet1
        .Shapes("Chart 1").Top = 32.25008
        .Shapes("Chart 1").Left = 14.25016
        .Shapes("Chart 1").Height = 126
        .Shapes("Chart 1").Width = 454.6053
        .Shapes("Chart 3").Top = 619.5
        .Shapes("Chart 3").Left = 13.5
        .Shapes("Chart 3").Height = 141.0441
        .Shapes("Chart 3").Width = 237.6
        .Shapes("Chart 7").Top = 164.25
        .Shapes("Chart 7").Left = 14.24992
        .Shapes("Chart 7").Height = 141.0441
        .Shapes("Chart 7").Width = 452.25
        .Shapes("Chart 8").Top = 313.3676
        .Shapes("Chart 8").Left = 14.24992
        .Shapes("Chart 8").Height = 140.1617
        .Shapes("Chart 8").Width = 452.25
        .Shapes("Chart 9").Top = 463.3676
        .Shapes("Chart 9").Left = 14.25
        .Shapes("Chart 9").Height = 140.1617
        .Shapes("Chart 9").Width = 237.6
        .Shapes("Salesperson").Top = 157.5
        .Shapes("Salesperson").Left = 483.0002
        .Shapes("Salesperson").Height = 203.7353
        .Shapes("Salesperson").Width = 101.2499
        .Shapes("Region 2").Top = 33.00008
        .Shapes("Region 2").Left = 482.2502
        .Shapes("Region 2").Height = 116.9999
        .Shapes("Region 2").Width = 101.2499
        .Shapes("TextBox 12").Top = 6
        .Shapes("TextBox 12").Left = 10.5
        .Shapes("TextBox 12").Height = 27
        .Shapes("TextBox 12").Width = 237
        .Shapes("TextBox 13").Top = 7.5
        .Shapes("TextBox 13").Left = 369.75
        .Shapes("TextBox 13").Height = 24
        .Shapes("TextBox 13").Width = 119.8235
        .Shapes("Picture 17").Top = 13.49992
        .Shapes("Picture 17").Left = 358.4999
        .Shapes("Picture 17").Height = 17.25
        .Shapes("Picture 17").Width = 17.25
    End With

End Sub

Sub View_Narrow()
'Window Width: 855.75
'Window Height: 498

    With Sheet1
        .Shapes("Chart 1").Top = 32.25008
        .Shapes("Chart 1").Left = 14.25024
        .Shapes("Chart 1").Height = 126
        .Shapes("Chart 1").Width = 289.4999
        .Shapes("Chart 3").Top = 619.5
        .Shapes("Chart 3").Left = 13.5
        .Shapes("Chart 3").Height = 141.0441
        .Shapes("Chart 3").Width = 237.6
        .Shapes("Chart 7").Top = 164.25
        .Shapes("Chart 7").Left = 14.25
        .Shapes("Chart 7").Height = 141.0441
        .Shapes("Chart 7").Width = 288
        .Shapes("Chart 8").Top = 313.3676
        .Shapes("Chart 8").Left = 14.25
        .Shapes("Chart 8").Height = 140.1617
        .Shapes("Chart 8").Width = 288
        .Shapes("Chart 9").Top = 463.3676
        .Shapes("Chart 9").Left = 14.25
        .Shapes("Chart 9").Height = 140.1617
        .Shapes("Chart 9").Width = 237.6
        .Shapes("Salesperson").Top = 157.5
        .Shapes("Salesperson").Left = 315.0002
        .Shapes("Salesperson").Height = 203.7353
        .Shapes("Salesperson").Width = 101.2499
        .Shapes("Region 2").Top = 33.00008
        .Shapes("Region 2").Left = 314.2502
        .Shapes("Region 2").Height = 116.9999
        .Shapes("Region 2").Width = 101.2499
        .Shapes("TextBox 12").Top = 6
        .Shapes("TextBox 12").Left = 10.5
        .Shapes("TextBox 12").Height = 27
        .Shapes("TextBox 12").Width = 237
        .Shapes("TextBox 13").Top = 7.5
        .Shapes("TextBox 13").Left = 206.25
        .Shapes("TextBox 13").Height = 24
        .Shapes("TextBox 13").Width = 119.8235
        .Shapes("Picture 17").Top = 13.49992
        .Shapes("Picture 17").Left = 194.9999
        .Shapes("Picture 17").Height = 17.25
        .Shapes("Picture 17").Width = 17.25
    End With

End Sub
