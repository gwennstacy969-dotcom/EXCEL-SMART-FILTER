Attribute VB_Name = "modColumnFilter"
Option Explicit

' ==============================================================================
' TOOL NAME    : EXCEL SMART FILTER
' MODULE       : modColumnFilter (Modul Pengolahan Filter Kolom)
' DESCRIPTION  : Ekstraksi nilai unik dinamis, penanganan nilai (Kosong),
'                filter teks, pengurutan A-Z/Z-A, dan eksekusi AutoFilter native.
' ==============================================================================

Public Const BLANK_LABEL As String = "(Kosong)"

''' <summary>
''' Mengambil daftar seluruh nilai unik dari suatu kolom di worksheet.
''' Membaca nilai tampilan sel aktual dan menangani nilai kosong.
''' </summary>
Public Function GetColumnUniqueValues( _
    ByVal ws As Worksheet, _
    ByVal headerRow As Long, _
    ByVal lastRow As Long, _
    ByVal colIdx As Long) As Variant
    
    Dim dict As Object
    Dim r As Long
    Dim cellText As String
    Dim hasBlank As Boolean
    Dim arrResult() As String
    Dim i As Long, j As Long
    Dim tempStr As String
    Dim vData As Variant
    
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1 ' TextCompare (case-insensitive)
    hasBlank = False
    
    If lastRow <= headerRow Then
        GetColumnUniqueValues = Array()
        Exit Function
    End If
    
    ' Baca seluruh data kolom ke memory array sekaligus untuk efisiensi tinggi
    If lastRow - headerRow = 1 Then
        ReDim vData(1 To 1, 1 To 1)
        vData(1, 1) = ws.Cells(headerRow + 1, colIdx).Text
    Else
        vData = ws.Range(ws.Cells(headerRow + 1, colIdx), ws.Cells(lastRow, colIdx)).Value
    End If
    
    For r = 1 To UBound(vData, 1)
        If IsError(vData(r, 1)) Then
            cellText = "#ERROR"
        ElseIf IsEmpty(vData(r, 1)) Or IsNull(vData(r, 1)) Then
            cellText = vbNullString
        Else
            cellText = Trim$(CStr(vData(r, 1)))
        End If
        
        If Len(cellText) = 0 Then
            hasBlank = True
        Else
            If Not dict.Exists(cellText) Then
                dict.Add cellText, True
            End If
        End If
    Next r
    
    ' Konversi ke array terurut
    Dim keys As Variant
    keys = dict.Keys
    Dim totalItems As Long
    totalItems = dict.Count + IIf(hasBlank, 1, 0)
    
    If totalItems = 0 Then
        GetColumnUniqueValues = Array()
        Exit Function
    End If
    
    ReDim arrResult(0 To totalItems - 1)
    Dim idx As Long: idx = 0
    
    If hasBlank Then
        arrResult(idx) = BLANK_LABEL
        idx = idx + 1
    End If
    
    For i = 0 To dict.Count - 1
        arrResult(idx) = CStr(keys(i))
        idx = idx + 1
    Next i
    
    ' Urutkan array nilai secara alfabetis (A-Z), pertahankan (Kosong) di atas
    Dim startSort As Long
    startSort = IIf(hasBlank, 1, 0)
    
    For i = startSort To UBound(arrResult) - 1
        For j = i + 1 To UBound(arrResult)
            If StrComp(arrResult(i), arrResult(j), vbTextCompare) > 0 Then
                tempStr = arrResult(i)
                arrResult(i) = arrResult(j)
                arrResult(j) = tempStr
            End If
        Next j
    Next i
    
    GetColumnUniqueValues = arrResult
    Set dict = Nothing
End Function

''' <summary>
''' Menerapkan seluruh filter (Periode Tanggal dan Kriteria Kolom) menggunakan AutoFilter native Excel.
''' Tidak ada baris yang dihapus; baris yang tidak cocok hanya disembunyikan.
''' </summary>
Public Function ExecuteApplyFilters( _
    ByVal ws As Worksheet, _
    ByVal headerRow As Long, _
    ByVal lastRow As Long, _
    ByVal lastCol As Long, _
    ByVal pType As PeriodType, _
    ByVal dateColIdx As Long, _
    ByVal dtStart As Date, _
    ByVal dtEnd As Date, _
    ByVal dictColFilters As Object) As Long
    
    Dim tableRange As Range
    Dim colKey As Variant
    Dim filterInfo As Object
    Dim mode As String
    Dim cIdx As Long
    Dim arrSelected As Variant
    Dim arrFilterValues() As String
    Dim i As Long, countVal As Long
    Dim hasKosong As Boolean
    
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Set tableRange = ws.Range(ws.Cells(headerRow, 1), ws.Cells(lastRow, lastCol))
    
    ' Pastikan AutoFilter aktif pada header
    If Not ws.AutoFilterMode Then
        tableRange.AutoFilter
    Else
        ' Jika sudah aktif, reset data tanpa menghilangkan dropdown
        If ws.FilterMode Then
            ws.ShowAllData
        End If
    End If
    
    ' 1. TERAPKAN FILTER PERIODE TANGGAL
    If pType <> Period_SemuaTanggal And dateColIdx > 0 And dateColIdx <= lastCol Then
        Dim dStartSerial As Double
        Dim dEndBoundarySerial As Double
        
        ' Int(dtStart) = 00:00:00 hari mulai
        dStartSerial = CDbl(DateSerial(Year(dtStart), Month(dtStart), Day(dtStart)))
        
        ' Int(dtEnd) + 1 = 00:00:00 hari berikutnya, sehingga seluruh data pada tanggal akhir (termasuk jam 23:59:59) masuk!
        dEndBoundarySerial = CDbl(DateSerial(Year(dtEnd), Month(dtEnd), Day(dtEnd)) + 1)
        
        tableRange.AutoFilter _
            Field:=dateColIdx, _
            Criteria1:=">=" & dStartSerial, _
            Operator:=xlAnd, _
            Criteria2:="<" & dEndBoundarySerial
    End If
    
    ' 2. TERAPKAN FILTER KOLOM DINAMIS
    If Not dictColFilters Is Nothing Then
        For Each colKey In dictColFilters.Keys
            Set filterInfo = dictColFilters(colKey)
            cIdx = CLng(colKey)
            mode = CStr(filterInfo("Mode"))
            
            If mode = "VALUES" Then
                arrSelected = filterInfo("SelectedValues")
                If IsArray(arrSelected) Then
                    If UBound(arrSelected) >= 0 Then
                        ' Periksa apakah ada nilai kosong
                        countVal = 0
                        hasKosong = False
                        For i = LBound(arrSelected) To UBound(arrSelected)
                            If CStr(arrSelected(i)) = BLANK_LABEL Then
                                hasKosong = True
                            Else
                                countVal = countVal + 1
                            End If
                        Next i
                        
                        ' Bangun kriteria AutoFilter
                        Dim totalCriteria As Long
                        totalCriteria = countVal + IIf(hasKosong, 1, 0)
                        
                        If totalCriteria > 0 Then
                            ReDim arrFilterValues(0 To totalCriteria - 1)
                            Dim vIdx As Long: vIdx = 0
                            
                            If hasKosong Then
                                arrFilterValues(vIdx) = "=" ' Excel AutoFilter mengenali "=" sebagai sel kosong
                                vIdx = vIdx + 1
                            End If
                            
                            For i = LBound(arrSelected) To UBound(arrSelected)
                                If CStr(arrSelected(i)) <> BLANK_LABEL Then
                                    arrFilterValues(vIdx) = CStr(arrSelected(i))
                                    vIdx = vIdx + 1
                                End If
                            Next i
                            
                            ' Terapkan AutoFilter multi-nilai native
                            tableRange.AutoFilter _
                                Field:=cIdx, _
                                Criteria1:=arrFilterValues, _
                                Operator:=xlFilterValues
                        End If
                    End If
                End If
                
            ElseIf mode = "TEXT" Then
                Dim strOp As String
                Dim strVal As String
                Dim strCriteria As String
                
                strOp = CStr(filterInfo("Operator"))
                strVal = CStr(filterInfo("Criteria"))
                
                Select Case strOp
                    Case "Sama Dengan"
                        strCriteria = "=" & strVal
                    Case "Tidak Sama Dengan"
                        strCriteria = "<>" & strVal
                    Case "Berisi"
                        strCriteria = "=*" & strVal & "*"
                    Case "Tidak Berisi"
                        strCriteria = "<>*" & strVal & "*"
                    Case "Dimulai Dengan"
                        strCriteria = "=" & strVal & "*"
                    Case "Diakhiri Dengan"
                        strCriteria = "=*" & strVal
                    Case Else
                        strCriteria = "=*" & strVal & "*"
                End Select
                
                tableRange.AutoFilter Field:=cIdx, Criteria1:=strCriteria
            End If
        Next colKey
    End If
    
    ' Hitung jumlah baris yang cocok
    ExecuteApplyFilters = CountVisibleRows(ws, headerRow, lastRow)
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Function
    
ErrHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Terjadi kesalahan saat menerapkan filter:" & vbCrLf & Err.Description, vbCritical, "Error AutoFilter"
    ExecuteApplyFilters = -1
End Function

''' <summary>
''' Mengurutkan data (A-Z atau Z-A) pada kolom tertentu.
''' Mempertahankan formula, merge cell, dan formatting.
''' </summary>
Public Sub SortWorksheetColumn( _
    ByVal ws As Worksheet, _
    ByVal headerRow As Long, _
    ByVal lastRow As Long, _
    ByVal lastCol As Long, _
    ByVal colIdx As Long, _
    ByVal isAscending As Boolean)
    
    Dim dataRange As Range
    Dim sortKeyRange As Range
    Dim sortOrder As XlSortOrder
    
    On Error GoTo ErrHandler
    If lastRow <= headerRow Then Exit Sub
    
    Set dataRange = ws.Range(ws.Cells(headerRow + 1, 1), ws.Cells(lastRow, lastCol))
    Set sortKeyRange = ws.Range(ws.Cells(headerRow + 1, colIdx), ws.Cells(lastRow, colIdx))
    
    sortOrder = IIf(isAscending, xlAscending, xlDescending)
    
    ws.Sort.SortFields.Clear
    ws.Sort.SortFields.Add _
        Key:=sortKeyRange, _
        SortOn:=xlSortOnValues, _
        Order:=sortOrder, _
        DataOption:=xlSortNormal
        
    With ws.Sort
        .SetRange dataRange
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Exit Sub
ErrHandler:
    MsgBox "Gagal mengurutkan kolom: " & Err.Description, vbExclamation, "Error Pengurutan"
End Sub

''' <summary>
''' Menghapus seluruh filter aktif pada worksheet.
''' </summary>
Public Sub ResetAllFilters(ByVal ws As Worksheet)
    On Error Resume Next
    If ws.FilterMode Then
        ws.ShowAllData
    End If
End Sub
