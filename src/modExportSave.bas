Attribute VB_Name = "modExportSave"
Option Explicit

' ==============================================================================
' TOOL NAME    : EXCEL SMART FILTER
' MODULE       : modExportSave (Modul Penyimpanan Salinan Hasil)
' DESCRIPTION  : Menyimpan salinan hasil filter dengan format & ekstensi yang sesuai
'                tanpa mengubah file sumber, dan membuka folder di Windows Explorer.
' ==============================================================================

''' <summary>
''' Menyimpan workbook salinan kerja ke file tujuan baru pilihan pengguna.
''' Mempertahankan format asli, formula, merge cells, dan status filter.
''' </summary>
Public Function SaveFilteredResultCopy( _
    ByVal wbWorking As Workbook, _
    ByVal strOriginalPath As String, _
    ByVal totalRows As Long, _
    ByVal matchingRows As Long) As String
    
    Dim fso As Object
    Dim strBaseName As String
    Dim strExt As String
    Dim strDefaultName As String
    Dim strInitialFolder As String
    Dim strFilterPrompt As String
    Dim vSavePath As Variant
    Dim strTargetFile As String
    Dim currentExt As String
    
    On Error GoTo ErrHandler
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    strBaseName = fso.GetBaseName(strOriginalPath)
    strExt = LCase$(fso.GetExtensionName(strOriginalPath))
    strInitialFolder = fso.GetParentFolderName(strOriginalPath)
    
    strDefaultName = strBaseName & "_Filtered_" & Format(Now, "yyyymmdd") & "." & strExt
    
    ' Tentukan filter dialog berdasarkan ekstensi file sumber
    Select Case strExt
        Case "xlsx"
            strFilterPrompt = "Excel Workbook (*.xlsx), *.xlsx"
        Case "xlsm"
            strFilterPrompt = "Excel Macro-Enabled Workbook (*.xlsm), *.xlsm"
        Case "xls"
            strFilterPrompt = "Excel 97-2003 Workbook (*.xls), *.xls"
        Case "xlsb"
            strFilterPrompt = "Excel Binary Workbook (*.xlsb), *.xlsb"
        Case Else
            strFilterPrompt = "Excel Workbook (*." & strExt & "), *." & strExt
    End Select
    
    ' Simpan perubahan sementara ke working copy terlebih dahulu
    wbWorking.Save
    
    ' Buka dialog Save As
    vSavePath = Application.GetSaveAsFilename( _
        InitialFileName:=strInitialFolder & "\" & strDefaultName, _
        FileFilter:=strFilterPrompt & ", Semua File (*.*), *.*", _
        Title:="Simpan Hasil Filter Excel (Salinan Baru)")
        
    If VarType(vSavePath) = vbBoolean Then
        ' Pengguna menekan Batal (Cancel)
        SaveFilteredResultCopy = vbNullString
        Exit Function
    End If
    
    strTargetFile = CStr(vSavePath)
    
    ' Pastikan tidak menimpa file sumber secara langsung
    If LCase$(strTargetFile) = LCase$(strOriginalPath) Then
        MsgBox "PERINGATAN KEAMANAN:" & vbCrLf & _
               "Anda tidak diperbolehkan menimpa file sumber asli!" & vbCrLf & _
               "Pilih nama file yang berbeda untuk menjaga integritas file sumber.", _
               vbCritical, "Penyimpanan Ditolak"
        SaveFilteredResultCopy = vbNullString
        Exit Function
    End If
    
    ' Simpan salinan hasil menggunakan SaveCopyAs
    wbWorking.SaveCopyAs strTargetFile
    
    ' Verifikasi bahwa file hasil benar-benar berhasil dibuat di disk
    If fso.FileExists(strTargetFile) Then
        If fso.GetFile(strTargetFile).Size > 0 Then
            g_lastSavedPath = strTargetFile
            SaveFilteredResultCopy = strTargetFile
            
            MsgBox "Salinan hasil filter berhasil disimpan ke:" & vbCrLf & _
                   strTargetFile & vbCrLf & vbCrLf & _
                   "Ringkasan Data:" & vbCrLf & _
                   "• Total Baris Sumber : " & Format(totalRows, "#,##0") & " baris" & vbCrLf & _
                   "• Baris Cocok Filter : " & Format(matchingRows, "#,##0") & " baris (" & _
                   Format(IIf(totalRows > 0, matchingRows / totalRows, 0), "0.0%") & ")" & vbCrLf & _
                   "• Format & Formula   : Utuh terjaga" & vbCrLf & _
                   "• File Sumber Asli   : Aman & tidak berubah", _
                   vbInformation, "Penyimpanan Berhasil"
        Else
            MsgBox "Penyimpanan gagal: File yang terbentuk memiliki ukuran 0 byte.", vbCritical, "Gagal Simpan"
            SaveFilteredResultCopy = vbNullString
        End If
    Else
        MsgBox "Penyimpanan gagal: File hasil tidak dapat diverifikasi pada direktori tujuan.", vbCritical, "Gagal Simpan"
        SaveFilteredResultCopy = vbNullString
    End If
    
    Set fso = Nothing
    Exit Function
    
ErrHandler:
    MsgBox "Terjadi kesalahan saat menyimpan file hasil:" & vbCrLf & Err.Description, vbCritical, "Error Simpan File"
    SaveFilteredResultCopy = vbNullString
End Function

''' <summary>
''' Membuka folder di Windows Explorer dan memilih file yang baru saja disimpan.
''' </summary>
Public Sub OpenFolderWithFileSelected(ByVal strFilePath As String)
    Dim fso As Object
    On Error GoTo ErrHandler
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Len(strFilePath) = 0 Or Not fso.FileExists(strFilePath) Then
        ' Jika belum ada file yang disimpan, buka folder file sumber
        If Len(g_strOriginalFile) > 0 And fso.FileExists(g_strOriginalFile) Then
            strFilePath = g_strOriginalFile
        Else
            MsgBox "Belum ada file yang dipilih atau disimpan.", vbInformation, "Info Folder"
            Exit Sub
        End If
    End If
    
    ' Buka Windows Explorer dan pilih (highlight) file tersebut
    Shell "explorer.exe /select,""" & strFilePath & """", vbNormalFocus
    Set fso = Nothing
    Exit Sub
    
ErrHandler:
    MsgBox "Gagal membuka folder di Windows Explorer: " & Err.Description, vbExclamation, "Error Explorer"
End Sub
