Attribute VB_Name = "modMain"
Option Explicit

' ==============================================================================
' TOOL NAME    : EXCEL SMART FILTER
' MODULE       : modMain (Modul Utama & Pengelolaan File)
' DESCRIPTION  : Membuka file sumber, membuat salinan kerja (working copy),
'                mendeteksi header, mendeteksi kolom tanggal, dan menghitung baris.
' ==============================================================================

Public g_strOriginalFile As String
Public g_strWorkingCopy As String
Public g_wbWorking As Workbook
Public g_wsWorking As Worksheet
Public g_headerRow As Long
Public g_dateColIdx As Long
Public g_lastRow As Long
Public g_lastCol As Long
Public g_lastSavedPath As String

''' <summary>
''' Entry point untuk menjalankan tools EXCEL SMART FILTER.
''' Dapat dijalankan melalui Alt+F8 atau tombol shortcut.
''' </summary>
Public Sub ShowSmartFilter()
    On Error GoTo ErrHandler
    frmSmartFilter.Show vbModeless
    Exit Sub
ErrHandler:
    MsgBox "Terjadi kesalahan saat membuka EXCEL SMART FILTER:" & vbCrLf & _
           Err.Description, vbCritical, "Error Peluncuran Tools"
End Sub

''' <summary>
''' Membuka dialog pemilihan file Excel (XLS, XLSX, XLSM, XLSB).
''' </summary>
Public Function SelectSourceFile() As String
    Dim fd As FileDialog
    On Error GoTo ErrHandler
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Pilih File Excel Bulanan (EXCEL SMART FILTER)"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "File Microsoft Excel (*.xlsx; *.xlsm; *.xls; *.xlsb)", "*.xlsx;*.xlsm;*.xls;*.xlsb"
        .Filters.Add "Semua File (*.*)", "*.*"
        .FilterIndex = 1
        
        If .Show = -1 Then
            SelectSourceFile = .SelectedItems(1)
        Else
            SelectSourceFile = vbNullString
        End If
    End With
    Exit Function
ErrHandler:
    MsgBox "Gagal membuka dialog file: " & Err.Description, vbExclamation, "Error Dialog"
    SelectSourceFile = vbNullString
End Function

''' <summary>
''' Membuat salinan kerja sementara (Working Copy) di folder TEMP.
''' Memastikan file sumber 100% aman, tidak terkunci, dan tidak berubah.
''' </summary>
Public Function CreateWorkingCopy(ByVal strSourcePath As String) As String
    Dim fso As Object
    Dim strTempFolder As String
    Dim strExt As String
    Dim strBaseName As String
    Dim strDestPath As String
    Dim strTimestamp As String
    
    On Error GoTo ErrHandler
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FileExists(strSourcePath) Then
        MsgBox "File sumber tidak ditemukan di lokasi:" & vbCrLf & strSourcePath, vbCritical, "File Tidak Ditemukan"
        CreateWorkingCopy = vbNullString
        Exit Function
    End If
    
    strTempFolder = Environ$("TEMP")
    If Len(strTempFolder) = 0 Then strTempFolder = fso.GetSpecialFolder(2) ' 2 = TemporaryFolder
    
    strExt = fso.GetExtensionName(strSourcePath)
    strBaseName = fso.GetBaseName(strSourcePath)
    strTimestamp = Format(Now, "yyyymmdd_hhnnss")
    
    strDestPath = strTempFolder & "\ESF_WorkCopy_" & strTimestamp & "_" & strBaseName & "." & strExt
    
    ' Salin file sumber ke lokasi sementara (tanpa membuka file sumber dalam mode tulis)
    fso.CopyFile strSourcePath, strDestPath, True
    
    CreateWorkingCopy = strDestPath
    Set fso = Nothing
    Exit Function
    
ErrHandler:
    MsgBox "Gagal membuat salinan kerja dari file sumber:" & vbCrLf & Err.Description & _
           vbCrLf & vbCrLf & "Pastikan file tidak sedang terkunci secara eksklusif oleh proses lain.", _
           vbCritical, "Error Salinan Kerja"
    CreateWorkingCopy = vbNullString
End Function

''' <summary>
''' Membuka workbook salinan kerja.
''' </summary>
Public Function OpenWorkingWorkbook(ByVal strPath As String) As Workbook
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    Set OpenWorkingWorkbook = Workbooks.Open( _
        Filename:=strPath, _
        UpdateLinks:=0, _
        ReadOnly:=False, _
        AddToMru:=False)
        
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Function
ErrHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "Gagal membuka salinan kerja: " & Err.Description, vbCritical, "Error Buka Workbook"
    Set OpenWorkingWorkbook = Nothing
End Function

''' <summary>
''' Mendeteksi nomor baris header secara otomatis.
''' Mencari baris pertama yang memiliki >= 3 sel berisi teks yang tidak kosong.
''' </summary>
Public Function DetectHeaderRow(ByVal ws As Worksheet, Optional ByVal maxScanRows As Long = 20) As Long
    Dim r As Long, c As Long
    Dim countNonEmpty As Long
    Dim cellText As String
    Dim maxColsToCheck As Long
    
    On Error Resume Next
    maxColsToCheck = ws.UsedRange.Columns.Count
    If maxColsToCheck > 50 Then maxColsToCheck = 50
    If maxColsToCheck < 2 Then maxColsToCheck = 10
    
    For r = 1 To maxScanRows
        countNonEmpty = 0
        For c = 1 To maxColsToCheck
            cellText = Trim$(CStr(ws.Cells(r, c).Text))
            If Len(cellText) > 0 Then
                ' Jika bukan sel hasil merge panjang atau banner tunggal
                countNonEmpty = countNonEmpty + 1
            End If
        Next c
        
        ' Baris header umumnya memiliki beberapa kolom teks sekaligus
        If countNonEmpty >= 3 Then
            DetectHeaderRow = r
            Exit Function
        End If
    Next r
    
    ' Default ke baris 1 jika tidak terdeteksi
    DetectHeaderRow = 1
End Function

''' <summary>
''' Mendeteksi kolom yang kemungkinan berisi tanggal (berdasarkan nama header atau format isi).
''' </summary>
Public Function DetectDateColumn(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal totalCols As Long) As Long
    Dim c As Long, r As Long
    Dim hText As String
    Dim sampleCell As Range
    Dim dateKeywords As Variant
    Dim kw As Variant
    Dim dateScores() As Long
    Dim maxScore As Long
    Dim bestCol As Long
    
    On Error Resume Next
    dateKeywords = Array("tanggal", "tgl", "date", "waktu", "time", "periode", "period", "transaksi", "created", "tgl_", "trx_date")
    
    ReDim dateScores(1 To totalCols)
    maxScore = 0
    bestCol = 1
    
    For c = 1 To totalCols
        hText = LCase$(Trim$(CStr(ws.Cells(headerRow, c).Text)))
        
        ' Cek kemiripan nama header
        For Each kw In dateKeywords
            If InStr(1, hText, CStr(kw), vbTextCompare) > 0 Then
                dateScores(c) = dateScores(c) + 10
                Exit For
            End If
        Next kw
        
        ' Cek sampel isi baris data (3 baris setelah header)
        For r = headerRow + 1 To headerRow + 5
            Set sampleCell = ws.Cells(r, c)
            If Not sampleCell Is Nothing Then
                If IsDate(sampleCell.Value) Then
                    dateScores(c) = dateScores(c) + 5
                ElseIf InStr(1, sampleCell.NumberFormat, "yy", vbTextCompare) > 0 Or _
                       InStr(1, sampleCell.NumberFormat, "mm", vbTextCompare) > 0 Or _
                       InStr(1, sampleCell.NumberFormat, "dd", vbTextCompare) > 0 Then
                    dateScores(c) = dateScores(c) + 4
                End If
            End If
        Next r
        
        If dateScores(c) > maxScore Then
            maxScore = dateScores(c)
            bestCol = c
        End If
    Next c
    
    If maxScore > 0 Then
        DetectDateColumn = bestCol
    Else
        DetectDateColumn = 1
    End If
End Function

''' <summary>
''' Menghitung jumlah baris data yang terlihat (visible) setelah filter diterapkan.
''' </summary>
Public Function CountVisibleRows(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal lastRow As Long) As Long
    Dim r As Long
    Dim visCount As Long
    
    If lastRow <= headerRow Then
        CountVisibleRows = 0
        Exit Function
    End If
    
    visCount = 0
    For r = headerRow + 1 To lastRow
        If Not ws.Rows(r).Hidden Then
            visCount = visCount + 1
        End If
    Next r
    
    CountVisibleRows = visCount
End Function

''' <summary>
''' Fungsi utilitas nama file dasar.
''' </summary>
Public Function GetFileBaseName(ByVal fullPath As String) As String
    Dim fso As Object
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    GetFileBaseName = fso.GetBaseName(fullPath)
    Set fso = Nothing
End Function

''' <summary>
''' Fungsi utilitas ekstensi file.
''' </summary>
Public Function GetFileExtension(ByVal fullPath As String) As String
    Dim fso As Object
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    GetFileExtension = fso.GetExtensionName(fullPath)
    Set fso = Nothing
End Function

''' <summary>
''' Fungsi utilitas folder induk dari path file.
''' </summary>
Public Function GetFolderPath(ByVal fullPath As String) As String
    Dim fso As Object
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    GetFolderPath = fso.GetParentFolderName(fullPath)
    Set fso = Nothing
End Function

''' <summary>
''' Menghapus file temporary dengan aman.
''' </summary>
Public Sub CleanupTempFile(ByVal tempPath As String)
    Dim fso As Object
    On Error Resume Next
    If Len(tempPath) > 0 Then
        Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.FileExists(tempPath) Then
            fso.DeleteFile tempPath, True
        End If
        Set fso = Nothing
    End If
End Sub
