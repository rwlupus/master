Attribute VB_Name = "progress_bar"
'Option Explicit

Sub start()

Dim dPoczatek As Double, dKoniec As Double, dCzas As Double
    Dim counter As Long, PctDone As Double
    Dim RowMax As Long, ColMax As Long, col As Long, row As Long
    
    Dim LCounts As Long
    
    RowMax = 200
    ColMax = 200


    LCounts = RowMax * ColMax
    Dim StartTime As Double
    Dim ETA As String



    StartTime = Now

    counter = 1
    Application.ScreenUpdating = False
    UserForm1.Show vbModeless
    
    
For col = 1 To ColMax
        For row = 1 To RowMax
          If Cells(row, col).Interior.ColorIndex <> 1 Then
          Cells(row, col) = WorksheetFunction.RandBetween(1, 9999)
          End If
          counter = counter + 1
        Next row
    Application.Wait Now + TimeValue("00:00:0" & WorksheetFunction.RandBetween(20, 60))

    PctDone = counter / LCounts

     ETA = Format((Now() - StartTime) / PctDone + StartTime, "dddd hh:mm")
     
    RefreshProgresBar counter, LCounts, PctDone, ETA

    
        
Next col
    Unload UserForm1
Application.ScreenUpdating = True



End Sub

Sub RefreshProgresBar(counter As Long, counts As Long, PctDone As Double, ETA As String)
        With UserForm1
            .FrameProgress.Caption = Space(20) & Format(PctDone, "0 %")
            .LabelProgress.Width = PctDone * (.FrameProgress.Width - 10)
            .Caption = counter & " of " & counts & " ETA: " & ETA
        End With

    DoEvents
End Sub
