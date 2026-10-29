VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSmartFilter 
   Caption         =   "EXCEL SMART FILTER - Professional Data Filtering Tool"
   ClientHeight    =   10200
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15000
   StartUpPosition =   1  'CenterOwner
   TypeLibVersion  =   32
End
Attribute VB_Name = "frmSmartFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ==============================================================================
' TOOL NAME    : EXCEL SMART FILTER
' FORM         : frmSmartFilter (UserForm Utama)
' DESCRIPTION  : Antarmuka utama untuk pemilihan file, worksheet, periode tanggal,
'                filter kolom dinamis (checkbox, search, text filter, sort),
'                penerapan filter native, dan penyimpanan salinan hasil.
' ==============================================================================

' Koleksi event handler kontrol dinamis
Private m_colEvents As Collection

' Dictionary kriteria filter kolom aktif: Key = ColIndex (Long), Value = Dictionary kriteria
Private m_dictColFilters As Object

' Cache nilai unik kolom yang sedang aktif: Key = ColIndex, Value = Array of Strings
Private m_dictColUniqueCache As Object

' Variabel kontrol form
Private txtFilePath As MSForms.TextBox
Private cboWorksheet As MSForms.ComboBox
Private txtHeaderRow As MSForms.TextBox
Private cboDateCol As MSForms.ComboBox

Private optAllDates As MSForms.OptionButton
Private optDaily As MSForms.OptionButton
Private optWeekly As MSForms.OptionButton
Private optDateRange As MSForms.OptionButton

Private lblDate1 As MSForms.Label
Private txtDate1 As MSForms.TextBox
Private lblDate2 As MSForms.Label
Private txtDate2 As MSForms.TextBox
Private lblPeriodPreview As MSForms.Label

Private cboColumnSelect As MSForms.ComboBox
Private txtSearch As MSForms.TextBox
Private lstValues As MSForms.ListBox

Private cboTextOp As MSForms.ComboBox
Private txtTextVal As MSForms.TextBox

Private txtSummary As MSForms.TextBox
Private lblStatus As MSForms.Label

' Status internal
Private m_bIsUpdatingList As Boolean
Private m_totalRows As Long
Private m_matchingRows As Long
Private m_currentColIdx As Long

Private Sub UserForm_Initialize()
    On Error GoTo ErrHandler
    
    Set m_colEvents = New Collection
    Set m_dictColFilters = CreateObject("Scripting.Dictionary")
    Set m_dictColUniqueCache = CreateObject("Scripting.Dictionary")
    
    m_bIsUpdatingList = False
    m_totalRows = 0
    m_matchingRows = 0
    m_currentColIdx = 0
    
    ' Ukuran Form Utama (Points)
    Me.Width = 760
    Me.Height = 560
    Me.BackColor = &HF4F6F8 ' Abu-abu terang modern
    
    ' Bangun antarmuka kontrol secara terstruktur
    BuildUserInterface
    
    ' Set nilai default periode
    optAllDates.Value = True
    txtDate1.Text = Format(Date, "yyyy-mm-dd")
    txtDate2.Text = Format(Date, "yyyy-mm-dd")
    UpdatePeriodUI
    
    Exit Sub
ErrHandler:
    MsgBox "Inisialisasi Form gagal: " & Err.Description, vbCritical, "Error Form"
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    On Error Resume Next
    ' Bersihkan workbook salinan kerja jika dibuka
    If Not g_wbWorking Is Nothing Then
        g_wbWorking.Close SaveChanges:=False
        Set g_wbWorking = Nothing
    End If
    If Len(g_strWorkingCopy) > 0 Then
        CleanupTempFile g_strWorkingCopy
    End If
End Sub

' ==============================================================================
' PEMBANGUNAN TAMPILAN USER INTERFACE
' ==============================================================================
Private Sub BuildUserInterface()
    ' 1. HEADER BANNER
    Dim lblTitle As MSForms.Label
    Set lblTitle = Me.Controls.Add("Forms.Label.1", "lblTitle", True)
    With lblTitle
        .Left = 12: .Top = 10: .Width = 730: .Height = 24
        .Caption = "EXCEL SMART FILTER"
        .Font.Name = "Segoe UI": .Font.Size = 14: .Font.Bold = True
        .ForeColor = &H803300 ' Corporate Navy
        .BackStyle = fmBackStyleTransparent
    End With
    
    Dim lblSub As MSForms.Label
    Set lblSub = Me.Controls.Add("Forms.Label.1", "lblSub", True)
    With lblSub
        .Left = 14: .Top = 32: .Width = 728: .Height = 16
        .Caption = "Filter Fleksibel Data Excel Bulanan • 100% Menjaga Format, Formula & Integritas File Sumber"
        .Font.Name = "Segoe UI": .Font.Size = 8.5: .ForeColor = &H666666
        .BackStyle = fmBackStyleTransparent
    End With
    
    ' Garis pembatas atas
    Dim lineSep As MSForms.Label
    Set lineSep = Me.Controls.Add("Forms.Label.1", "lineSep", True)
    With lineSep
        .Left = 12: .Top = 50: .Width = 726: .Height = 1
        .BackColor = &HD0D7DE
    End With
    
    ' 2. FRAME 1: FILE & WORKSHEET (Top Full Width)
    Dim fraFile As MSForms.Frame
    Set fraFile = Me.Controls.Add("Forms.Frame.1", "fraFile", True)
    With fraFile
        .Left = 12: .Top = 56: .Width = 726: .Height = 68
        .Caption = " 1. File Excel & Lembar Kerja (Sumber) "
        .Font.Name = "Segoe UI": .Font.Bold = True: .Font.Size = 8.5
        .ForeColor = &H803300: .BackColor = &HF4F6F8
    End With
    
    AddLabel fraFile, "File:", 10, 16, 30, 14
    Set txtFilePath = AddTextBox(fraFile, "txtFilePath", 42, 14, 460, 20, True)
    txtFilePath.Text = "Belum ada file dipilih..."
    txtFilePath.ForeColor = &H777777
    
    Dim btnBrowse As MSForms.CommandButton
    Set btnBrowse = AddButton(fraFile, "btnBrowse", "Pilih File...", 510, 14, 85, 20, "BTN_BROWSE")
    
    Dim btnReload As MSForms.CommandButton
    Set btnReload = AddButton(fraFile, "btnReload", "Muat Sheet", 602, 14, 110, 20, "BTN_RELOAD_SHEET")
    
    AddLabel fraFile, "Pilih Worksheet:", 10, 40, 85, 14
    Set cboWorksheet = AddComboBox(fraFile, "cboWorksheet", 98, 38, 200, 20, "CBO_WORKSHEET")
    
    AddLabel fraFile, "Baris Header:", 320, 40, 75, 14
    Set txtHeaderRow = AddTextBox(fraFile, "txtHeaderRow", 398, 38, 45, 20, False, "TXT_HEADER_ROW")
    txtHeaderRow.Text = "1"
    
    ' 3. KOLOM KIRI (LEBAR 358): FILTER PERIODE TANGGAL
    Dim fraPeriod As MSForms.Frame
    Set fraPeriod = Me.Controls.Add("Forms.Frame.1", "fraPeriod", True)
    With fraPeriod
        .Left = 12: .Top = 130: .Width = 358: .Height = 225
        .Caption = " 2. Filter Periode Tanggal "
        .Font.Name = "Segoe UI": .Font.Bold = True: .Font.Size = 8.5
        .ForeColor = &H803300: .BackColor = &HF4F6F8
    End With
    
    AddLabel fraPeriod, "Kolom Tanggal:", 10, 18, 85, 14
    Set cboDateCol = AddComboBox(fraPeriod, "cboDateCol", 98, 16, 245, 20, "CBO_DATE_COL")
    
    Set optAllDates = AddOptionButton(fraPeriod, "optAllDates", "Semua Tanggal (Tanpa Filter Tanggal)", 12, 42, 330, 18, "OPT_ALL_DATES")
    Set optDaily = AddOptionButton(fraPeriod, "optDaily", "Per Hari (Tanggal Tertentu)", 12, 62, 330, 18, "OPT_DAILY")
    Set optWeekly = AddOptionButton(fraPeriod, "optWeekly", "Per Minggu (Senin - Minggu Otomatis)", 12, 82, 330, 18, "OPT_WEEKLY")
    Set optDateRange = AddOptionButton(fraPeriod, "optDateRange", "Rentang Tanggal Bebas (Mulai s/d Akhir)", 12, 102, 330, 18, "OPT_DATE_RANGE")
    
    Set lblDate1 = AddLabel(fraPeriod, "Tanggal:", 12, 126, 85, 14)
    Set txtDate1 = AddTextBox(fraPeriod, "txtDate1", 100, 124, 110, 20, False, "TXT_DATE1")
    
    Set lblDate2 = AddLabel(fraPeriod, "Tanggal Akhir:", 12, 150, 85, 14)
    Set txtDate2 = AddTextBox(fraPeriod, "txtDate2", 100, 148, 110, 20, False, "TXT_DATE2")
    
    ' Tombol pintasan tanggal
    AddButton fraPeriod, "btnToday", "Hari Ini", 220, 124, 60, 20, "BTN_TODAY")
    AddButton fraPeriod, "btnMonthStart", "Awal Bln", 285, 124, 60, 20, "BTN_MONTH_START")
    AddButton fraPeriod, "btnMonthEnd", "Akhir Bln", 220, 148, 60, 20, "BTN_MONTH_END")
    AddButton fraPeriod, "btnPreviewPeriod", "Hitung", 285, 148, 60, 20, "BTN_CALC_PERIOD")
    
    Set lblPeriodPreview = AddLabel(fraPeriod, "Status Periode: Semua Tanggal", 12, 176, 332, 38)
    lblPeriodPreview.Font.Size = 7.5: lblPeriodPreview.ForeColor = &H993300
    
    ' 4. KOLOM KANAN (LEBAR 358): FILTER KOLOM DINAMIS
    Dim fraCol As MSForms.Frame
    Set fraCol = Me.Controls.Add("Forms.Frame.1", "fraCol", True)
    With fraCol
        .Left = 380: .Top = 130: .Width = 358: .Height = 225
        .Caption = " 3. Filter Kolom Dinamis & Nilai "
        .Font.Name = "Segoe UI": .Font.Bold = True: .Font.Size = 8.5
        .ForeColor = &H803300: .BackColor = &HF4F6F8
    End With
    
    AddLabel fraCol, "Pilih Kolom:", 8, 18, 65, 14
    Set cboColumnSelect = AddComboBox(fraCol, "cboColumnSelect", 75, 16, 175, 20, "CBO_COL_SELECT")
    
    AddButton fraCol, "btnSortAZ", "A-Z", 255, 16, 42, 20, "BTN_SORT_AZ")
    AddButton fraCol, "btnSortZA", "Z-A", 302, 16, 42, 20, "BTN_SORT_ZA")
    
    AddLabel fraCol, "Cari Nilai:", 8, 42, 65, 14
    Set txtSearch = AddTextBox(fraCol, "txtSearch", 75, 40, 269, 20, False, "TXT_SEARCH")
    
    ' ListBox MultiSelect dengan Checkbox
    Set lstValues = fraCol.Controls.Add("Forms.ListBox.1", "lstValues", True)
    With lstValues
        .Left = 8: .Top = 64: .Width = 336: .Height = 100
        .ListStyle = fmListStyleOption     ' Tampilkan Checkbox
        .MultiSelect = fmMultiSelectMulti  ' Multi-pilihan
        .Font.Name = "Segoe UI": .Font.Size = 8.5
    End With
    RegisterEvent lstValues, "LST_VALUES", "lst"
    
    ' Tombol Aksi Kolom
    AddButton fraCol, "btnSaveCol", "Simpan Kriteria Kolom", 8, 168, 145, 22, "BTN_SAVE_COL")
    AddButton fraCol, "btnClearCol", "Hapus Filter Kolom Ini", 160, 168, 135, 22, "BTN_CLEAR_COL")
    
    ' Baris Filter Teks Alternatif
    AddLabel fraCol, "Filter Teks:", 8, 198, 60, 14
    Set cboTextOp = AddComboBox(fraCol, "cboTextOp", 68, 196, 105, 20)
    With cboTextOp
        .AddItem "Sama Dengan"
        .AddItem "Tidak Sama Dengan"
        .AddItem "Berisi"
        .AddItem "Tidak Berisi"
        .AddItem "Dimulai Dengan"
        .AddItem "Diakhiri Dengan"
        .ListIndex = 2 ' Default: Berisi
    End With
    Set txtTextVal = AddTextBox(fraCol, "txtTextVal", 178, 196, 95, 20)
    AddButton fraCol, "btnApplyText", "Terapkan", 278, 196, 66, 20, "BTN_APPLY_TEXT")
    
    ' 5. RINGKASAN FILTER & STATUS DATA (Bawah)
    Dim fraSummary As MSForms.Frame
    Set fraSummary = Me.Controls.Add("Forms.Frame.1", "fraSummary", True)
    With fraSummary
        .Left = 12: .Top = 360: .Width = 726: .Height = 110
        .Caption = " 4. Ringkasan Kriteria Aktif & Status Baris Data "
        .Font.Name = "Segoe UI": .Font.Bold = True: .Font.Size = 8.5
        .ForeColor = &H803300: .BackColor = &HF4F6F8
    End With
    
    Set txtSummary = fraSummary.Controls.Add("Forms.TextBox.1", "txtSummary", True)
    With txtSummary
        .Left = 10: .Top = 18: .Width = 705: .Height = 62
        .MultiLine = True: .ScrollBars = fmScrollBarsVertical
        .Locked = True: .BackColor = &HFFFFFF
        .Font.Name = "Consolas": .Font.Size = 8.5
        .Text = "(Belum ada filter kriteria yang diterapkan)"
    End With
    
    Set lblStatus = AddLabel(fraSummary, "Total Data: 0 baris | Cocok: 0 baris (0.0%)", 10, 85, 705, 18)
    lblStatus.Font.Name = "Segoe UI": lblStatus.Font.Bold = True: lblStatus.Font.Size = 9
    lblStatus.ForeColor = &H2A52BE
    
    ' 6. ACTION BAR TOMBOL UTAMA (Bottom Bar)
    Dim btnApplyAll As MSForms.CommandButton
    Set btnApplyAll = AddButton(Me, "btnApplyAll", "TERAPKAN FILTER", 12, 480, 150, 32, "BTN_APPLY_ALL")
    btnApplyAll.Font.Bold = True: btnApplyAll.BackColor = &H803300: btnApplyAll.ForeColor = &HFFFFFF
    
    Dim btnClearAll As MSForms.CommandButton
    Set btnClearAll = AddButton(Me, "btnClearAll", "Hapus Semua Filter", 168, 480, 125, 32, "BTN_CLEAR_ALL")
    
    Dim btnViewSheet As MSForms.CommandButton
    Set btnViewSheet = AddButton(Me, "btnViewSheet", "Lihat Hasil di Excel", 298, 480, 125, 32, "BTN_VIEW_SHEET")
    
    Dim btnSaveExcel As MSForms.CommandButton
    Set btnSaveExcel = AddButton(Me, "btnSaveExcel", "Simpan Hasil Excel", 428, 480, 135, 32, "BTN_SAVE_EXCEL")
    btnSaveExcel.Font.Bold = True
    
    Dim btnOpenFolder As MSForms.CommandButton
    Set btnOpenFolder = AddButton(Me, "btnOpenFolder", "Buka Folder Hasil", 568, 480, 110, 32, "BTN_OPEN_FOLDER")
    
    Dim btnClose As MSForms.CommandButton
    Set btnClose = AddButton(Me, "btnClose", "Tutup", 683, 480, 55, 32, "BTN_CLOSE")
End Sub

' ==============================================================================
' EVENT DISPATCHERS & HANDLERS
' ==============================================================================
Public Sub OnCustomButtonClick(ByVal eventKey As String)
    On Error GoTo ErrHandler
    Select Case eventKey
        Case "BTN_BROWSE": DoBrowseFile
        Case "BTN_RELOAD_SHEET": DoLoadWorksheet
        Case "BTN_TODAY": SetQuickDate 0
        Case "BTN_MONTH_START": SetQuickDate 1
        Case "BTN_MONTH_END": SetQuickDate 2
        Case "BTN_CALC_PERIOD": UpdatePeriodUI
        Case "BTN_SORT_AZ": DoSortSelectedColumn True
        Case "BTN_SORT_ZA": DoSortSelectedColumn False
        Case "BTN_SAVE_COL": DoSaveCurrentColumnFilter
        Case "BTN_CLEAR_COL": DoClearCurrentColumnFilter
        Case "BTN_APPLY_TEXT": DoApplyTextFilterCurrentColumn
        Case "BTN_APPLY_ALL": DoApplyAllFilters
        Case "BTN_CLEAR_ALL": DoClearAllFilters
        Case "BTN_VIEW_SHEET": DoViewResultsInExcel
        Case "BTN_SAVE_EXCEL": DoSaveFilteredResults
        Case "BTN_OPEN_FOLDER": DoOpenOutputFolder
        Case "BTN_CLOSE": Unload Me
    End Select
    Exit Sub
ErrHandler:
    MsgBox "Error Tombol (" & eventKey & "): " & Err.Description, vbCritical, "Error"
End Sub

Public Sub OnCustomOptionClick(ByVal eventKey As String)
    On Error Resume Next
    UpdatePeriodUI
End Sub

Public Sub OnCustomComboChange(ByVal eventKey As String)
    On Error Resume Next
    Select Case eventKey
        Case "CBO_WORKSHEET": DoLoadWorksheet
        Case "CBO_COL_SELECT": OnColumnSelectChanged
        Case "CBO_DATE_COL": UpdatePeriodUI
    End Select
End Sub

Public Sub OnCustomListChange(ByVal eventKey As String)
    On Error Resume Next
    If eventKey = "LST_VALUES" Then
        OnListValuesItemChecked
    End If
End Sub

Public Sub OnCustomTextChange(ByVal eventKey As String)
    On Error Resume Next
    Select Case eventKey
        Case "TXT_SEARCH": DoFilterListBySearch
        Case "TXT_DATE1", "TXT_DATE2": UpdatePeriodUI
    End Select
End Sub

' ==============================================================================
' LOGIKA PROSES & FITUR UTAMA
' ==============================================================================
Private Sub DoBrowseFile()
    Dim strFile As String
    strFile = SelectSourceFile()
    If Len(strFile) = 0 Then Exit Sub
    
    g_strOriginalFile = strFile
    txtFilePath.Text = strFile
    txtFilePath.ForeColor = &H000000
    
    ' Buat salinan kerja di folder TEMP agar file sumber 100% aman
    If Not g_wbWorking Is Nothing Then
        g_wbWorking.Close SaveChanges:=False
        Set g_wbWorking = Nothing
    End If
    If Len(g_strWorkingCopy) > 0 Then CleanupTempFile g_strWorkingCopy
    
    g_strWorkingCopy = CreateWorkingCopy(strFile)
    If Len(g_strWorkingCopy) = 0 Then Exit Sub
    
    Set g_wbWorking = OpenWorkingWorkbook(g_strWorkingCopy)
    If g_wbWorking Is Nothing Then Exit Sub
    
    ' Isi daftar worksheet yang tersedia
    cboWorksheet.Clear
    Dim ws As Worksheet
    For Each ws In g_wbWorking.Worksheets
        If ws.Visible = xlSheetVisible Then
            cboWorksheet.AddItem ws.Name
        End If
    Next ws
    
    If cboWorksheet.ListCount > 0 Then
        cboWorksheet.ListIndex = 0
    End If
End Sub

Private Sub DoLoadWorksheet()
    Dim wsName As String
    Dim hRow As Long
    Dim c As Long
    Dim colName As String
    
    If g_wbWorking Is Nothing Then Exit Sub
    If cboWorksheet.ListIndex < 0 Then Exit Sub
    
    wsName = cboWorksheet.Text
    Set g_wsWorking = g_wbWorking.Worksheets(wsName)
    
    ' Deteksi baris header secara otomatis jika belum diisi user
    If IsNumeric(txtHeaderRow.Text) And Val(txtHeaderRow.Text) > 0 Then
        hRow = Val(txtHeaderRow.Text)
    Else
        hRow = DetectHeaderRow(g_wsWorking)
        txtHeaderRow.Text = CStr(hRow)
    End If
    g_headerRow = hRow
    
    ' Tentukan batas kolom dan baris
    g_lastCol = g_wsWorking.Cells(g_headerRow, g_wsWorking.Columns.Count).End(xlToLeft).Column
    If g_lastCol < 1 Then g_lastCol = 1
    
    g_lastRow = g_wsWorking.Cells(g_wsWorking.Rows.Count, 1).End(xlUp).Row
    If g_lastRow <= g_headerRow Then
        g_lastRow = g_wsWorking.UsedRange.Rows.Count + g_wsWorking.UsedRange.Row - 1
    End If
    
    m_totalRows = IIf(g_lastRow > g_headerRow, g_lastRow - g_headerRow, 0)
    m_matchingRows = m_totalRows
    
    ' Reset filter cache & kriteria
    m_dictColFilters.RemoveAll
    m_dictColUniqueCache.RemoveAll
    
    ' Isi ComboBox Kolom Tanggal & ComboBox Kolom Pilihan
    cboDateCol.Clear
    cboColumnSelect.Clear
    
    For c = 1 To g_lastCol
        colName = Trim$(CStr(g_wsWorking.Cells(g_headerRow, c).Text))
        If Len(colName) = 0 Then
            colName = "[Kolom " & Split(g_wsWorking.Cells(1, c).Address, "$")(1) & "]"
        End If
        cboDateCol.AddItem colName
        cboColumnSelect.AddItem colName
    Next c
    
    ' Deteksi kolom tanggal otomatis
    g_dateColIdx = DetectDateColumn(g_wsWorking, g_headerRow, g_lastCol)
    If g_dateColIdx >= 1 And g_dateColIdx <= cboDateCol.ListCount Then
        cboDateCol.ListIndex = g_dateColIdx - 1
    End If
    
    If cboColumnSelect.ListCount > 0 Then
        cboColumnSelect.ListIndex = 0
    End If
    
    UpdateStatusLabel
    UpdateSummaryText
End Sub

Private Sub OnColumnSelectChanged()
    Dim colIdx As Long
    If cboColumnSelect.ListIndex < 0 Then Exit Sub
    If g_wsWorking Is Nothing Then Exit Sub
    
    colIdx = cboColumnSelect.ListIndex + 1
    m_currentColIdx = colIdx
    
    ' Ambil daftar nilai unik dari cache atau langsung dari worksheet
    Dim arrValues As Variant
    If Not m_dictColUniqueCache.Exists(colIdx) Then
        arrValues = GetColumnUniqueValues(g_wsWorking, g_headerRow, g_lastRow, colIdx)
        m_dictColUniqueCache.Add colIdx, arrValues
    Else
        arrValues = m_dictColUniqueCache(colIdx)
    End If
    
    PopulateListBoxValues arrValues, vbNullString
End Sub

Private Sub PopulateListBoxValues(ByVal arrValues As Variant, ByVal strFilterKeyword As String)
    Dim i As Long
    Dim valStr As String
    Dim isMatch As Boolean
    Dim previouslySelected As Object
    
    m_bIsUpdatingList = True
    lstValues.Clear
    
    If Not IsArray(arrValues) Then
        m_bIsUpdatingList = False
        Exit Sub
    End If
    If UBound(arrValues) < LBound(arrValues) Then
        m_bIsUpdatingList = False
        Exit Sub
    End If
    
    ' Ambil kriteria tersimpan sebelumnya jika ada
    Set previouslySelected = CreateObject("Scripting.Dictionary")
    previouslySelected.CompareMode = 1
    
    If m_dictColFilters.Exists(m_currentColIdx) Then
        Dim filterInfo As Object
        Set filterInfo = m_dictColFilters(m_currentColIdx)
        If filterInfo("Mode") = "VALUES" Then
            Dim arrSaved As Variant
            arrSaved = filterInfo("SelectedValues")
            If IsArray(arrSaved) Then
                For i = LBound(arrSaved) To UBound(arrSaved)
                    previouslySelected(CStr(arrSaved(i))) = True
                Next i
            End If
        End If
    End If
    
    ' Item 0: (Pilih Semua)
    lstValues.AddItem "(Pilih Semua)"
    lstValues.Selected(0) = True
    
    For i = LBound(arrValues) To UBound(arrValues)
        valStr = CStr(arrValues(i))
        isMatch = True
        
        If Len(strFilterKeyword) > 0 Then
            If InStr(1, valStr, strFilterKeyword, vbTextCompare) = 0 Then
                isMatch = False
            End If
        End If
        
        If isMatch Then
            lstValues.AddItem valStr
            Dim newIdx As Long
            newIdx = lstValues.ListCount - 1
            
            If previouslySelected.Count > 0 Then
                lstValues.Selected(newIdx) = previouslySelected.Exists(valStr)
            Else
                lstValues.Selected(newIdx) = True
            End If
        End If
    Next i
    
    m_bIsUpdatingList = False
    Set previouslySelected = Nothing
End Sub

Private Sub OnListValuesItemChecked()
    If m_bIsUpdatingList Then Exit Sub
    If lstValues.ListCount = 0 Then Exit Sub
    
    ' Tangani klik pada item 0 "(Pilih Semua)"
    Dim i As Long
    Static bInSelectAll As Boolean
    If bInSelectAll Then Exit Sub
    
    ' Cek apakah item 0 yang diubah
    If lstValues.ListIndex = 0 Then
        bInSelectAll = True
        Dim stateSelectAll As Boolean
        stateSelectAll = lstValues.Selected(0)
        
        m_bIsUpdatingList = True
        For i = 1 To lstValues.ListCount - 1
            lstValues.Selected(i) = stateSelectAll
        Next i
        m_bIsUpdatingList = False
        bInSelectAll = False
    End If
End Sub

Private Sub DoFilterListBySearch()
    If m_currentColIdx <= 0 Then Exit Sub
    If Not m_dictColUniqueCache.Exists(m_currentColIdx) Then Exit Sub
    
    Dim kw As String
    kw = Trim$(txtSearch.Text)
    PopulateListBoxValues m_dictColUniqueCache(m_currentColIdx), kw
End Sub

Private Sub DoSaveCurrentColumnFilter()
    If m_currentColIdx <= 0 Then Exit Sub
    
    Dim arrSelected() As String
    Dim countSelected As Long
    Dim i As Long
    Dim filterInfo As Object
    
    countSelected = 0
    For i = 1 To lstValues.ListCount - 1
        If lstValues.Selected(i) Then
            countSelected = countSelected + 1
        End If
    Next i
    
    If countSelected = 0 Then
        MsgBox "Tidak ada nilai yang dicentang. Minimal pilih satu nilai atau klik Hapus Filter.", vbInformation, "Info"
        Exit Sub
    End If
    
    ReDim arrSelected(0 To countSelected - 1)
    Dim sIdx As Long: sIdx = 0
    For i = 1 To lstValues.ListCount - 1
        If lstValues.Selected(i) Then
            arrSelected(sIdx) = lstValues.List(i)
            sIdx = sIdx + 1
        End If
    Next i
    
    Set filterInfo = CreateObject("Scripting.Dictionary")
    filterInfo("ColIdx") = m_currentColIdx
    filterInfo("ColName") = cboColumnSelect.Text
    filterInfo("Mode") = "VALUES"
    filterInfo("SelectedValues") = arrSelected
    filterInfo("TotalAvailable") = lstValues.ListCount - 1
    
    If m_dictColFilters.Exists(m_currentColIdx) Then
        Set m_dictColFilters(m_currentColIdx) = filterInfo
    Else
        m_dictColFilters.Add m_currentColIdx, filterInfo
    End If
    
    UpdateSummaryText
    MsgBox "Kriteria untuk kolom '" & cboColumnSelect.Text & "' berhasil disimpan.", vbInformation, "Tersimpan"
End Sub

Private Sub DoClearCurrentColumnFilter()
    If m_currentColIdx <= 0 Then Exit Sub
    If m_dictColFilters.Exists(m_currentColIdx) Then
        m_dictColFilters.Remove m_currentColIdx
    End If
    
    ' Centang kembali semua nilai pada listbox
    m_bIsUpdatingList = True
    Dim i As Long
    For i = 0 To lstValues.ListCount - 1
        lstValues.Selected(i) = True
    Next i
    m_bIsUpdatingList = False
    
    UpdateSummaryText
End Sub

Private Sub DoApplyTextFilterCurrentColumn()
    If m_currentColIdx <= 0 Then Exit Sub
    Dim strVal As String
    strVal = Trim$(txtTextVal.Text)
    
    If Len(strVal) = 0 Then
        MsgBox "Masukkan kata kunci filter teks terlebih dahulu.", vbExclamation, "Filter Teks"
        Exit Sub
    End If
    
    Dim filterInfo As Object
    Set filterInfo = CreateObject("Scripting.Dictionary")
    filterInfo("ColIdx") = m_currentColIdx
    filterInfo("ColName") = cboColumnSelect.Text
    filterInfo("Mode") = "TEXT"
    filterInfo("Operator") = cboTextOp.Text
    filterInfo("Criteria") = strVal
    
    If m_dictColFilters.Exists(m_currentColIdx) Then
        Set m_dictColFilters(m_currentColIdx) = filterInfo
    Else
        m_dictColFilters.Add m_currentColIdx, filterInfo
    End If
    
    UpdateSummaryText
    MsgBox "Filter teks untuk kolom '" & cboColumnSelect.Text & "' berhasil diset (" & _
           cboTextOp.Text & " '" & strVal & "').", vbInformation, "Filter Teks"
End Sub

Private Sub DoSortSelectedColumn(ByVal isAscending As Boolean)
    If g_wsWorking Is Nothing Then Exit Sub
    If m_currentColIdx <= 0 Then Exit Sub
    
    SortWorksheetColumn g_wsWorking, g_headerRow, g_lastRow, g_lastCol, m_currentColIdx, isAscending
    MsgBox "Kolom '" & cboColumnSelect.Text & "' berhasil diurutkan " & _
           IIf(isAscending, "A ke Z.", "Z ke A."), vbInformation, "Pengurutan Berhasil"
End Sub

Private Sub UpdatePeriodUI()
    Dim pType As PeriodType
    Dim strPrompt As String
    Dim dtStart As Date, dtEnd As Date
    
    pType = GetCurrentPeriodType()
    
    Select Case pType
        Case Period_SemuaTanggal
            lblDate1.Visible = False: txtDate1.Visible = False
            lblDate2.Visible = False: txtDate2.Visible = False
            lblPeriodPreview.Caption = "Periode Aktif: Semua Tanggal (Seluruh data ditampilkan)"
            
        Case Period_PerHari
            lblDate1.Visible = True: txtDate1.Visible = True
            lblDate1.Caption = "Pilih Tanggal:"
            lblDate2.Visible = False: txtDate2.Visible = False
            If GetPeriodDateRange(pType, txtDate1.Text, "", dtStart, dtEnd, strPrompt) Then
                lblPeriodPreview.Caption = strPrompt
            End If
            
        Case Period_PerMinggu
            lblDate1.Visible = True: txtDate1.Visible = True
            lblDate1.Caption = "Tanggal Acuan:"
            lblDate2.Visible = False: txtDate2.Visible = False
            If GetPeriodDateRange(pType, txtDate1.Text, "", dtStart, dtEnd, strPrompt) Then
                lblPeriodPreview.Caption = strPrompt
            End If
            
        Case Period_RentangTanggal
            lblDate1.Visible = True: txtDate1.Visible = True
            lblDate1.Caption = "Tanggal Mulai:"
            lblDate2.Visible = True: txtDate2.Visible = True
            lblDate2.Caption = "Tanggal Akhir:"
            If GetPeriodDateRange(pType, txtDate1.Text, txtDate2.Text, dtStart, dtEnd, strPrompt) Then
                lblPeriodPreview.Caption = strPrompt
            End If
    End Select
End Sub

Private Function GetCurrentPeriodType() As PeriodType
    If optDaily.Value Then
        GetCurrentPeriodType = Period_PerHari
    ElseIf optWeekly.Value Then
        GetCurrentPeriodType = Period_PerMinggu
    ElseIf optDateRange.Value Then
        GetCurrentPeriodType = Period_RentangTanggal
    Else
        GetCurrentPeriodType = Period_SemuaTanggal
    End If
End Function

Private Sub SetQuickDate(ByVal quickMode As Integer)
    Dim dNow As Date: dNow = Date
    Select Case quickMode
        Case 0 ' Hari ini
            txtDate1.Text = Format(dNow, "yyyy-mm-dd")
            txtDate2.Text = Format(dNow, "yyyy-mm-dd")
        Case 1 ' Awal bulan
            txtDate1.Text = Format(DateSerial(Year(dNow), Month(dNow), 1), "yyyy-mm-dd")
        Case 2 ' Akhir bulan
            txtDate2.Text = Format(DateSerial(Year(dNow), Month(dNow) + 1, 0), "yyyy-mm-dd")
    End Select
    UpdatePeriodUI
End Sub

Private Sub DoApplyAllFilters()
    If g_wsWorking Is Nothing Then
        MsgBox "Silakan pilih file dan muat worksheet terlebih dahulu.", vbExclamation, "Peringatan"
        Exit Sub
    End If
    
    Dim pType As PeriodType
    Dim dtStart As Date, dtEnd As Date
    Dim strPeriodLabel As String
    Dim dateColIdx As Long
    
    pType = GetCurrentPeriodType()
    
    ' Validasi Tanggal jika menggunakan periode
    If pType <> Period_SemuaTanggal Then
        If cboDateCol.ListIndex < 0 Then
            MsgBox "Silakan pilih Kolom Tanggal yang digunakan.", vbExclamation, "Validasi Tanggal"
            Exit Sub
        End If
        dateColIdx = cboDateCol.ListIndex + 1
        
        If Not GetPeriodDateRange(pType, txtDate1.Text, txtDate2.Text, dtStart, dtEnd, strPeriodLabel) Then
            Exit Sub
        End If
    Else
        dateColIdx = 0
    End If
    
    ' Terapkan seluruh filter native
    Dim resultCount As Long
    resultCount = ExecuteApplyFilters(g_wsWorking, g_headerRow, g_lastRow, g_lastCol, _
                                     pType, dateColIdx, dtStart, dtEnd, m_dictColFilters)
                                     
    If resultCount >= 0 Then
        m_matchingRows = resultCount
        UpdateStatusLabel
        UpdateSummaryText
        
        If m_matchingRows = 0 Then
            MsgBox "Filter selesai diterapkan, tetapi tidak ada baris data yang cocok dengan seluruh kriteria yang dipilih." & _
                   vbCrLf & "Data tidak dihapus, hanya disembunyikan.", vbInformation, "Hasil: 0 Baris Cocok"
        Else
            MsgBox "Filter berhasil diterapkan!" & vbCrLf & _
                   "Total baris cocok: " & Format(m_matchingRows, "#,##0") & " dari " & _
                   Format(m_totalRows, "#,##0") & " baris (" & _
                   Format(m_matchingRows / IIf(m_totalRows > 0, m_totalRows, 1), "0.0%") & ").", _
                   vbInformation, "Filter Berhasil"
        End If
    End If
End Sub

Private Sub DoClearAllFilters()
    If g_wsWorking Is Nothing Then Exit Sub
    
    ResetAllFilters g_wsWorking
    m_dictColFilters.RemoveAll
    optAllDates.Value = True
    UpdatePeriodUI
    
    m_matchingRows = m_totalRows
    UpdateStatusLabel
    UpdateSummaryText
    
    If m_currentColIdx > 0 Then
        OnColumnSelectChanged
    End If
    
    MsgBox "Seluruh filter telah dibersihkan. Semua baris data ditampilkan kembali.", vbInformation, "Filter Direset"
End Sub

Private Sub DoViewResultsInExcel()
    If g_wsWorking Is Nothing Then Exit Sub
    
    ' Tampilkan workbook di depan
    Application.Visible = True
    g_wbWorking.Activate
    g_wsWorking.Activate
    
    ' Beri pesan singkat kepada pengguna
    MsgBox "Anda sekarang dapat meninjau hasil filter langsung pada Microsoft Excel." & vbCrLf & _
           "Jendela tools EXCEL SMART FILTER tetap aktif dan siap digunakan kembali kapan saja.", _
           vbInformation, "Lihat Hasil Excel"
End Sub

Private Sub DoSaveFilteredResults()
    If g_wsWorking Is Nothing Then
        MsgBox "Belum ada file atau hasil filter yang dapat disimpan.", vbExclamation, "Peringatan"
        Exit Sub
    End If
    
    SaveFilteredResultCopy g_wbWorking, g_strOriginalFile, m_totalRows, m_matchingRows
End Sub

Private Sub DoOpenOutputFolder()
    OpenFolderWithFileSelected g_lastSavedPath
End Sub

Private Sub UpdateStatusLabel()
    Dim pct As Double
    If m_totalRows > 0 Then
        pct = (m_matchingRows / m_totalRows) * 100#
    Else
        pct = 0#
    End If
    
    lblStatus.Caption = "Total Data: " & Format(m_totalRows, "#,##0") & " baris  |  " & _
                        "Cocok dengan Filter: " & Format(m_matchingRows, "#,##0") & " baris (" & _
                        Format(pct, "0.0") & "%)"
End Sub

Private Sub UpdateSummaryText()
    Dim s As String
    Dim colKey As Variant
    Dim fInfo As Object
    Dim pType As PeriodType
    Dim dtStart As Date, dtEnd As Date
    Dim pLabel As String
    
    pType = GetCurrentPeriodType()
    s = "[PERIODE TANGGAL]" & vbCrLf
    If pType = Period_SemuaTanggal Then
        s = s & "  • Semua Tanggal (Tanpa Batasan)" & vbCrLf
    Else
        GetPeriodDateRange pType, txtDate1.Text, txtDate2.Text, dtStart, dtEnd, pLabel
        s = s & "  • Kolom  : " & cboDateCol.Text & vbCrLf
        s = s & "  • Rentang: " & pLabel & vbCrLf
    End If
    
    s = s & vbCrLf & "[KRITERIA FILTER KOLOM (Logika: AND Antar Kolom)]" & vbCrLf
    If m_dictColFilters.Count = 0 Then
        s = s & "  • (Tidak ada kriteria kolom aktif - seluruh kolom ditampilkan)" & vbCrLf
    Else
        For Each colKey In m_dictColFilters.Keys
            Set fInfo = m_dictColFilters(colKey)
            s = s & "  • " & fInfo("ColName") & " : "
            If fInfo("Mode") = "VALUES" Then
                Dim arrVal As Variant: arrVal = fInfo("SelectedValues")
                Dim countV As Long: countV = UBound(arrVal) - LBound(arrVal) + 1
                Dim totalV As Long: totalV = fInfo("TotalAvailable")
                s = s & countV & " dari " & totalV & " nilai dipilih (" & Join(arrVal, ", ") & ")" & vbCrLf
            Else
                s = s & fInfo("Operator") & " '" & fInfo("Criteria") & "'" & vbCrLf
            End If
        Next colKey
    End If
    
    txtSummary.Text = s
End Sub

' ==============================================================================
' HELPER GENERATOR KONTROL & REGISTRASI EVENT
' ==============================================================================
Private Function AddLabel(ByVal container As Object, ByVal txt As String, _
                          ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single) As MSForms.Label
    Dim lbl As MSForms.Label
    Set lbl = container.Controls.Add("Forms.Label.1", "lbl_" & Replace(txt, " ", "_") & "_" & CStr(l), True)
    With lbl
        .Left = l: .Top = t: .Width = w: .Height = h
        .Caption = txt: .Font.Name = "Segoe UI": .Font.Size = 8.5
        .BackStyle = fmBackStyleTransparent
    End With
    Set AddLabel = lbl
End Function

Private Function AddTextBox(ByVal container As Object, ByVal nameStr As String, _
                            ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single, _
                            ByVal isLocked As Boolean, Optional ByVal eventKey As String = "") As MSForms.TextBox
    Dim txtBox As MSForms.TextBox
    Set txtBox = container.Controls.Add("Forms.TextBox.1", nameStr, True)
    With txtBox
        .Left = l: .Top = t: .Width = w: .Height = h
        .Locked = isLocked: .Font.Name = "Segoe UI": .Font.Size = 8.5
        If isLocked Then .BackColor = &HF0F0F0
    End With
    If Len(eventKey) > 0 Then RegisterEvent txtBox, eventKey, "txt"
    Set AddTextBox = txtBox
End Function

Private Function AddComboBox(ByVal container As Object, ByVal nameStr As String, _
                             ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single, _
                             Optional ByVal eventKey As String = "") As MSForms.ComboBox
    Dim cboBox As MSForms.ComboBox
    Set cboBox = container.Controls.Add("Forms.ComboBox.1", nameStr, True)
    With cboBox
        .Left = l: .Top = t: .Width = w: .Height = h
        .Font.Name = "Segoe UI": .Font.Size = 8.5: .Style = fmStyleDropDownList
    End With
    If Len(eventKey) > 0 Then RegisterEvent cboBox, eventKey, "cbo"
    Set AddComboBox = cboBox
End Function

Private Function AddOptionButton(ByVal container As Object, ByVal nameStr As String, ByVal capStr As String, _
                                 ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single, _
                                 ByVal eventKey As String) As MSForms.OptionButton
    Dim optBtn As MSForms.OptionButton
    Set optBtn = container.Controls.Add("Forms.OptionButton.1", nameStr, True)
    With optBtn
        .Left = l: .Top = t: .Width = w: .Height = h
        .Caption = capStr: .Font.Name = "Segoe UI": .Font.Size = 8.5
        .GroupName = "grpPeriod": .BackStyle = fmBackStyleTransparent
    End With
    RegisterEvent optBtn, eventKey, "opt"
    Set AddOptionButton = optBtn
End Function

Private Function AddButton(ByVal container As Object, ByVal nameStr As String, ByVal capStr As String, _
                           ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single, _
                           ByVal eventKey As String) As MSForms.CommandButton
    Dim cmdBtn As MSForms.CommandButton
    Set cmdBtn = container.Controls.Add("Forms.CommandButton.1", nameStr, True)
    With cmdBtn
        .Left = l: .Top = t: .Width = w: .Height = h
        .Caption = capStr: .Font.Name = "Segoe UI": .Font.Size = 8.5
    End With
    RegisterEvent cmdBtn, eventKey, "btn"
    Set AddButton = cmdBtn
End Function

Private Sub RegisterEvent(ByVal ctl As Object, ByVal eventKey As String, ByVal ctlType As String)
    Dim h As clsEventHandler
    Set h = New clsEventHandler
    h.EventKey = eventKey
    Set h.ParentForm = Me
    
    Select Case ctlType
        Case "btn": Set h.btn = ctl
        Case "opt": Set h.opt = ctl
        Case "cbo": Set h.cbo = ctl
        Case "lst": Set h.lst = ctl
        Case "txt": Set h.txt = ctl
    End Select
    
    m_colEvents.Add h
End Sub

