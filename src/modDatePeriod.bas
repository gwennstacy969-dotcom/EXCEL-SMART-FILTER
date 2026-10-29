Attribute VB_Name = "modDatePeriod"
Option Explicit

' ==============================================================================
' TOOL NAME    : EXCEL SMART FILTER
' MODULE       : modDatePeriod (Modul Perhitungan Periode Tanggal)
' DESCRIPTION  : Parsing tanggal fleksibel, perhitungan batas mingguan
'                (Senin - Minggu), validasi rentang tanggal, dan penanganan jam.
' ==============================================================================

Public Enum PeriodType
    Period_SemuaTanggal = 0
    Period_PerHari = 1
    Period_PerMinggu = 2
    Period_RentangTanggal = 3
End Enum

''' <summary>
''' Menghitung tanggal Senin dan Minggu berdasarkan tanggal acuan.
''' Menangani pergantian bulan dan pergantian tahun dengan tepat.
''' </summary>
Public Sub CalculateWeekBoundaries(ByVal dtRef As Date, ByRef dtMonday As Date, ByRef dtSunday As Date)
    Dim dayOfWeek As Integer
    ' vbMonday: 1 = Senin, 2 = Selasa, ..., 7 = Minggu
    dayOfWeek = Weekday(dtRef, vbMonday)
    dtMonday = DateAdd("d", -(dayOfWeek - 1), dtRef)
    dtSunday = DateAdd("d", 6, dtMonday)
End Sub

''' <summary>
''' Menghitung rentang tanggal aktif berdasarkan pilihan periode.
''' </summary>
Public Function GetPeriodDateRange( _
    ByVal pType As PeriodType, _
    ByVal strDateInput1 As String, _
    ByVal strDateInput2 As String, _
    ByRef dtStart As Date, _
    ByRef dtEnd As Date, _
    ByRef strPeriodLabel As String) As Boolean
    
    Dim d1 As Date, d2 As Date
    Dim dMon As Date, dSun As Date
    
    GetPeriodDateRange = False
    
    Select Case pType
        Case Period_SemuaTanggal
            dtStart = DateSerial(1900, 1, 1)
            dtEnd = DateSerial(9999, 12, 31)
            strPeriodLabel = "Semua Tanggal (Tanpa Batasan Periode)"
            GetPeriodDateRange = True
            Exit Function
            
        Case Period_PerHari
            If Not ParseDateFlexible(strDateInput1, d1) Then
                MsgBox "Format tanggal harian tidak valid: '" & strDateInput1 & "'." & vbCrLf & _
                       "Gunakan format: YYYY-MM-DD atau DD/MM/YYYY.", vbExclamation, "Validasi Tanggal"
                Exit Function
            End If
            dtStart = d1
            dtEnd = d1
            strPeriodLabel = "Per Hari: " & FormatDateIndo(d1)
            GetPeriodDateRange = True
            Exit Function
            
        Case Period_PerMinggu
            If Not ParseDateFlexible(strDateInput1, d1) Then
                MsgBox "Format tanggal acuan minggu tidak valid: '" & strDateInput1 & "'." & vbCrLf & _
                       "Gunakan format: YYYY-MM-DD atau DD/MM/YYYY.", vbExclamation, "Validasi Tanggal"
                Exit Function
            End If
            CalculateWeekBoundaries d1, dMon, dSun
            dtStart = dMon
            dtEnd = dSun
            strPeriodLabel = "Per Minggu (Senin - Minggu): " & vbCrLf & _
                             FormatDateIndo(dMon) & " s/d " & FormatDateIndo(dSun)
            GetPeriodDateRange = True
            Exit Function
            
        Case Period_RentangTanggal
            If Not ParseDateFlexible(strDateInput1, d1) Then
                MsgBox "Tanggal Mulai tidak valid: '" & strDateInput1 & "'.", vbExclamation, "Validasi Tanggal"
                Exit Function
            End If
            If Not ParseDateFlexible(strDateInput2, d2) Then
                MsgBox "Tanggal Akhir tidak valid: '" & strDateInput2 & "'.", vbExclamation, "Validasi Tanggal"
                Exit Function
            End If
            
            ' Validasi tanggal akhir tidak lebih awal dari tanggal mulai
            If d2 < d1 Then
                MsgBox "Tanggal Akhir (" & FormatDateIndo(d2) & ") tidak boleh lebih awal dari Tanggal Mulai (" & _
                       FormatDateIndo(d1) & ")!", vbExclamation, "Validasi Rentang Tanggal"
                Exit Function
            End If
            
            dtStart = d1
            dtEnd = d2
            strPeriodLabel = "Rentang Tanggal: " & FormatDateIndo(d1) & " s/d " & FormatDateIndo(d2)
            GetPeriodDateRange = True
            Exit Function
    End Select
End Function

''' <summary>
''' Mem-parsing berbagai format tanggal (termasuk yang memiliki jam dan nama bulan Indonesia).
''' </summary>
Public Function ParseDateFlexible(ByVal vValue As Variant, ByRef outDate As Date) As Boolean
    Dim strVal As String
    Dim parts As Variant
    Dim y As Long, m As Long, d As Long
    Dim spacePos As Long
    
    ParseDateFlexible = False
    If IsError(vValue) Then Exit Function
    If IsNull(vValue) Or IsEmpty(vValue) Then Exit Function
    
    ' Jika sudah berupa tipe Date asli
    If VarType(vValue) = vbDate Then
        outDate = DateSerial(Year(vValue), Month(vValue), Day(vValue))
        ParseDateFlexible = True
        Exit Function
    End If
    
    ' Jika berupa angka serial Excel
    If IsNumeric(vValue) Then
        Dim dblVal As Double
        dblVal = CDbl(vValue)
        If dblVal >= 1 And dblVal <= 2958465 Then ' Range tahun 1900-9999
            outDate = CDate(Int(dblVal))
            ParseDateFlexible = True
            Exit Function
        End If
    End If
    
    strVal = Trim$(CStr(vValue))
    If Len(strVal) = 0 Then Exit Function
    
    ' Hilangkan bagian jam jika ada (misal "2026-10-15 14:30:00")
    spacePos = InStr(1, strVal, " ")
    If spacePos > 1 Then
        strVal = Trim$(Left$(strVal, spacePos - 1))
    End If
    
    ' Cek pemisah "-"
    If InStr(1, strVal, "-") > 0 Then
        parts = Split(strVal, "-")
        If UBound(parts) = 2 Then
            ' YYYY-MM-DD
            If Len(parts(0)) = 4 And IsNumeric(parts(0)) And IsNumeric(parts(1)) And IsNumeric(parts(2)) Then
                y = CLng(parts(0)): m = CLng(parts(1)): d = CLng(parts(2))
                If ValidDateValues(y, m, d) Then
                    outDate = DateSerial(y, m, d)
                    ParseDateFlexible = True
                    Exit Function
                End If
            ' DD-MM-YYYY
            ElseIf Len(parts(2)) = 4 And IsNumeric(parts(2)) And IsNumeric(parts(1)) And IsNumeric(parts(0)) Then
                d = CLng(parts(0)): m = CLng(parts(1)): y = CLng(parts(2))
                If ValidDateValues(y, m, d) Then
                    outDate = DateSerial(y, m, d)
                    ParseDateFlexible = True
                    Exit Function
                End If
            End If
        End If
    End If
    
    ' Cek pemisah "/"
    If InStr(1, strVal, "/") > 0 Then
        parts = Split(strVal, "/")
        If UBound(parts) = 2 Then
            ' DD/MM/YYYY atau MM/DD/YYYY
            If Len(parts(2)) = 4 And IsNumeric(parts(2)) And IsNumeric(parts(0)) And IsNumeric(parts(1)) Then
                y = CLng(parts(2))
                ' Asumsikan DD/MM/YYYY terlebih dahulu
                d = CLng(parts(0)): m = CLng(parts(1))
                If m > 12 And d <= 12 Then
                    ' Swap jika m > 12 (MM/DD/YYYY)
                    Dim tmp As Long
                    tmp = d: d = m: m = tmp
                End If
                If ValidDateValues(y, m, d) Then
                    outDate = DateSerial(y, m, d)
                    ParseDateFlexible = True
                    Exit Function
                End If
            ' YYYY/MM/DD
            ElseIf Len(parts(0)) = 4 And IsNumeric(parts(0)) And IsNumeric(parts(1)) And IsNumeric(parts(2)) Then
                y = CLng(parts(0)): m = CLng(parts(1)): d = CLng(parts(2))
                If ValidDateValues(y, m, d) Then
                    outDate = DateSerial(y, m, d)
                    ParseDateFlexible = True
                    Exit Function
                End If
            End If
        End If
    End If
    
    ' Fallback ke built-in IsDate
    If IsDate(strVal) Then
        outDate = CDate(strVal)
        outDate = DateSerial(Year(outDate), Month(outDate), Day(outDate))
        ParseDateFlexible = True
        Exit Function
    End If
End Function

Private Function ValidDateValues(ByVal y As Long, ByVal m As Long, ByVal d As Long) As Boolean
    If y < 1900 Or y > 2100 Then Exit Function
    If m < 1 Or m > 12 Then Exit Function
    If d < 1 Or d > 31 Then Exit Function
    ValidDateValues = True
End Function

''' <summary>
''' Memformat tanggal ke teks deskriptif berbahasa Indonesia lengkap.
''' </summary>
Public Function FormatDateIndo(ByVal dt As Date) As String
    Dim arrHari As Variant
    Dim arrBulan As Variant
    Dim strHari As String
    Dim strBulan As String
    Dim dayOfWeek As Integer
    
    arrHari = Array("Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu")
    arrBulan = Array("", "Januari", "Februari", "Maret", "April", "Mei", "Juni", _
                        "Juli", "Agustus", "September", "Oktober", "November", "Desember")
                        
    dayOfWeek = Weekday(dt, vbSunday) - 1
    strHari = arrHari(dayOfWeek)
    strBulan = arrBulan(Month(dt))
    
    FormatDateIndo = strHari & ", " & Day(dt) & " " & strBulan & " " & Year(dt)
End Function
