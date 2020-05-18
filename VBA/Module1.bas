Attribute VB_Name = "Module1"
Option Explicit

Private Const HEADER As String = "date,company,ID,state"

Private m_bWrongHeader As Boolean
Private m_sHeader As String

Private m_bMissingDates As Boolean
Private m_avDataGaps() As Variant

Private m_bMultipleCompany As Boolean
Private m_avMultipleCompany() As Variant

Private m_bWrongDatesFormat As Boolean
Private m_avWrongDatesFormat() As Variant


Sub RunChecks()

    Dim avData() As Variant
    Dim strt As Double
    
    strt = Timer
    avData = Range("A1").CurrentRegion.Value2
    
'    m_bWrongHeader = WrongHeader(avData)
'
'    m_bMissingDates = MissingDates(avData)
'
'    m_bMultipleCompany = MultipleCompany(avData)

'Stop

    m_bWrongDatesFormat = WrongDatesFormat(Range("A1").CurrentRegion)

Debug.Print Timer - strt
End Sub

Private Function WrongHeader(p_avData() As Variant) As Boolean

    Dim i As Long
    
    ReDim avHeader(1 To UBound(p_avData, 2)) As Variant
    
    For i = LBound(p_avData, 2) To UBound(p_avData, 2)
        avHeader(i) = p_avData(1, i)
    Next
    
    
    m_sHeader = Join(avHeader, ",")
    If Not m_sHeader = HEADER Then WrongHeader = True

   Debug.Print "WrongHeader", WrongHeader

End Function


Private Function MissingDates(p_avData() As Variant) As Boolean

    Dim i As Long
    Dim lDataGap As Long
    
    Dim avUniqueDates() As Variant

    Dim dictDates As Object
    Set dictDates = CreateObject("Scripting.Dictionary")
    
    For i = LBound(p_avData, 1) + 1 To UBound(p_avData, 1)
        dictDates(p_avData(i, 1)) = Empty
    Next
    
    avUniqueDates = dictDates.Keys
    Quicksort avUniqueDates, LBound(avUniqueDates), UBound(avUniqueDates)
    
    ReDim m_avDataGaps(1 To 3, 1 To 1)
    m_avDataGaps(1, 1) = "Date"
    m_avDataGaps(2, 1) = "Next Date"
    m_avDataGaps(3, 1) = "Missing days"
    

    For i = LBound(avUniqueDates) + 1 To UBound(avUniqueDates)
    
        If CDate(avUniqueDates(i)) >= DateSerial(Year(Date), Month(Date), 1) Then Exit For
        
        If avUniqueDates(i) <> avUniqueDates(i - 1) + 1 Then
            
            ReDim Preserve m_avDataGaps(1 To 3, 1 To UBound(m_avDataGaps, 2) + 1)
            
            lDataGap = DateDiff("d", CDate(avUniqueDates(i - 1)), CDate(avUniqueDates(i))) - 1
            
            m_avDataGaps(1, UBound(m_avDataGaps, 2)) = CDate(avUniqueDates(i - 1))
            m_avDataGaps(2, UBound(m_avDataGaps, 2)) = CDate(avUniqueDates(i))
            m_avDataGaps(3, UBound(m_avDataGaps, 2)) = lDataGap
        
        End If
    
    Next
    
    If UBound(m_avDataGaps, 2) > 1 Then MissingDates = True
    Debug.Print "MissingDates", MissingDates

End Function

Private Function MultipleCompany(p_avData() As Variant) As Boolean

    Dim i As Long
    Dim lDataGap As Long
    
    Dim avUniqueCompanies() As Variant

    Dim dictCompanies As Object, dictMultipleCompanies As Object, dictInner As Object
    Set dictCompanies = CreateObject("Scripting.Dictionary")
    Set dictMultipleCompanies = CreateObject("Scripting.Dictionary")

    
    ReDim m_avMultipleCompany(1 To 2, 1 To 1)
    m_avMultipleCompany(1, 1) = "Company"
    m_avMultipleCompany(2, 1) = "States"
    
    For i = LBound(p_avData, 1) + 1 To UBound(p_avData, 1)

        If Not dictCompanies.Exists(p_avData(i, 2)) Then
            dictCompanies.Add p_avData(i, 2), p_avData(i, 4)
        Else
            If dictCompanies(p_avData(i, 2)) <> p_avData(i, 4) Then
                    If Not dictMultipleCompanies.Exists(p_avData(i, 2)) Then
                        Set dictInner = CreateObject("Scripting.Dictionary")
                        dictInner.Add dictCompanies(p_avData(i, 2)), Empty
                        dictInner.Add p_avData(i, 4), Empty
                        dictMultipleCompanies.Add p_avData(i, 2), dictInner
                        
                    Else
                        dictInner(p_avData(i, 4)) = Empty
                    End If
            End If
        End If

    Next
    

    
    If dictMultipleCompanies.Count > 0 Then
        
        MultipleCompany = True
        Dim cmp As Variant
        Dim avTemp As Variant
        Debug.Print "MultipleCompany", MultipleCompany
        
        For Each cmp In dictMultipleCompanies.Keys
            avTemp = dictMultipleCompanies(cmp).Keys
            Quicksort avTemp, LBound(avTemp), UBound(avTemp)
            Debug.Print cmp, Join(avTemp, ",")
        Next

    End If
    
    
    'm_avMultipleCompany = dictMultipleCompanies.Keys
    'Quicksort m_avMultipleCompany, LBound(m_avMultipleCompany), UBound(m_avMultipleCompany)
    
    
'    For i = 0 To dictMultipleCompanies.Count - 1
'    Debug.Print dictMultipleCompanies.Keys()(i), dictMultipleCompanies.Items()(i)
'    Next
 '   Stop




'    For i = LBound(avUniqueCompanies) + 1 To UBound(avUniqueCompanies)
'
'
'
'        If avUniqueCompanies(i) <> avUniqueCompanies(i - 1) + 1 Then
'
'            ReDim Preserve m_avDataGaps(1 To 3, 1 To UBound(m_avDataGaps, 2) + 1)
'
'            lDataGap = DateDiff("d", CDate(avUniqueCompanies(i - 1)), CDate(avUniqueCompanies(i))) - 1
'
'            m_avDataGaps(1, UBound(m_avDataGaps, 2)) = CDate(avUniqueDates(i - 1))
'            m_avDataGaps(2, UBound(m_avDataGaps, 2)) = CDate(avUniqueDates(i))
'            m_avDataGaps(3, UBound(m_avDataGaps, 2)) = lDataGap
'
'        End If
'
'    Next
'

    
End Function



Private Function WrongDatesFormat(p_rngDateRange As Range) As Boolean
    
    Dim avData() As String
    Dim i As Long
    Dim ccell As Range

   Stop
    p_rngDateRange.Offset(1, 0).Select
        For Each ccell In p_rngDateRange.Offset(1, 0)
        
                If ccell.NumberFormatLocal <> "dd.mm.rrrr" Then
                WrongDatesFormat = True
                Exit For
                
                End If
                
        Next
        
        Debug.Print "WrongDatesFormat", WrongDatesFormat, ccell.Address, ccell.NumberFormatLocal
    
End Function


Sub BubbleSort(MyArray() As Variant)
'Sorts a one-dimensional VBA array from smallest to largest
'using the bubble sort algorithm.
Dim i As Long, j As Long
Dim Temp As Variant
 
For i = LBound(MyArray) To UBound(MyArray) - 1
    For j = i + 1 To UBound(MyArray)
        If MyArray(i) > MyArray(j) Then
            Temp = MyArray(j)
            MyArray(j) = MyArray(i)
            MyArray(i) = Temp
        End If
    Next j
Next i
End Sub

End Sub

Sub Quicksort(vArray As Variant, arrLbound As Long, arrUbound As Long)
'Sorts a one-dimensional VBA array from smallest to largest
'using a very fast quicksort algorithm variant.
Dim pivotVal As Variant
Dim vSwap    As Variant
Dim tmpLow   As Long
Dim tmpHi    As Long
 
tmpLow = arrLbound
tmpHi = arrUbound
pivotVal = vArray((arrLbound + arrUbound) \ 2)
 
While (tmpLow <= tmpHi) 'divide
   While (vArray(tmpLow) < pivotVal And tmpLow < arrUbound)
      tmpLow = tmpLow + 1
   Wend
  
   While (pivotVal < vArray(tmpHi) And tmpHi > arrLbound)
      tmpHi = tmpHi - 1
   Wend
 
   If (tmpLow <= tmpHi) Then
      vSwap = vArray(tmpLow)
      vArray(tmpLow) = vArray(tmpHi)
      vArray(tmpHi) = vSwap
      tmpLow = tmpLow + 1
      tmpHi = tmpHi - 1
   End If
Wend
 
  If (arrLbound < tmpHi) Then Quicksort vArray, arrLbound, tmpHi 'conquer
  If (tmpLow < arrUbound) Then Quicksort vArray, tmpLow, arrUbound 'conquer
End Sub
