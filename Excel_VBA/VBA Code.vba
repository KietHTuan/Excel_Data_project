	'Module code standard'
	
	Option Explicit
	
	Dim NextRun As Date  ' Module-level variable for scheduling
	
	Sub ColorTreeMapBoxes_AGNA_Revised()
	    Dim sht As Worksheet
	    Dim chtObj As ChartObject
	    Dim cht As Chart
	    Dim srs As Series
	    Dim pt As Point
	    Dim pctChange As Double
	    Dim sectorName As String
	    Dim i As Long, dataRow As Long
	    Dim companyName As String
	
	    ' Set the worksheet
	    Set sht = ThisWorkbook.Sheets("Market_Data")
	
	    ' Loop through each ChartObject
	    For Each chtObj In sht.ChartObjects
	        Set cht = chtObj.Chart
	        If cht.ChartType = xlTreemap Then
	            sectorName = cht.ChartTitle.Text
	            Set srs = cht.SeriesCollection(1)
	
	            For i = 1 To srs.Points.Count
	                Set pt = srs.Points(i)
	                On Error Resume Next
	                companyName = pt.DataLabel.Text
	                On Error GoTo 0
	
	                If companyName <> "" Then
	                    For dataRow = 2 To sht.Cells(Rows.Count, "C").End(xlUp).Row
	                        If sht.Cells(dataRow, "C").Value = sectorName And sht.Cells(dataRow, "A").Value = companyName Then
	                            If IsNumeric(sht.Cells(dataRow, "F").Value) Then
	                                pctChange = sht.Cells(dataRow, "F").Value
	
	                                Select Case pctChange
	                                    Case Is < -0.05
	                                        pt.Format.Fill.ForeColor.RGB = RGB(255, 0, 0)
	                                    Case -0.05 To -0.02
	                                        pt.Format.Fill.ForeColor.RGB = RGB(200, 50, 50)
	                                    Case -0.02 To 0
	                                        pt.Format.Fill.ForeColor.RGB = RGB(150, 0, 0)
	                                    Case 0 To 0.02
	                                        pt.Format.Fill.ForeColor.RGB = RGB(0, 150, 0)
	                                    Case 0.02 To 0.05
	                                        pt.Format.Fill.ForeColor.RGB = RGB(50, 200, 50)
	                                    Case Is > 0.05
	                                        pt.Format.Fill.ForeColor.RGB = RGB(0, 255, 0)
	                                End Select
	                                Exit For
	                            Else
	                                Debug.Print "Skipping row " & dataRow & " (non-numeric %)."
	                            End If
	                        End If
	                    Next dataRow
	                Else
	                    Debug.Print "Missing company name for point " & i & " in " & sectorName
	                End If
	            Next i
	        End If
	    Next chtObj
	
	    Debug.Print "Update completed at " & Now
	    ScheduleNextRun  ' <-- Critical automation line added here
	End Sub
	
	Sub ScheduleNextRun()
	    NextRun = Now + TimeValue("00:05:00")  ' Updates every 5 minutes
	    Application.OnTime NextRun, "ColorTreeMapBoxes_AGNA_Revised"
	End Sub
	
	Sub StopAutomation()
	    On Error Resume Next
	    Application.OnTime NextRun, "ColorTreeMapBoxes_AGNA_Revised", , False
	    On Error GoTo 0
	    Debug.Print "Automation stopped at " & Now
	End Sub
	
	
	
'Paste this to this workbook code to run auto refresh '
	
	Private Sub Workbook_Open()
	    ColorTreeMapBoxes_AGNA_Revised  ' Starts automation when file opens
	End Sub
