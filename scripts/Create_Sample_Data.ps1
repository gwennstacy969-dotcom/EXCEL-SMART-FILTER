$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false

$wb = $xl.Workbooks.Add()

# Sheet 1: Transaksi_Bulanan
$ws1 = $wb.Worksheets.Item(1)
$ws1.Name = "Transaksi_Bulanan"

# Title Banner on Row 1
$banner = $ws1.Range("A1:L1")
$banner.Merge()
$ws1.Range("A1").Formula = "LAPORAN TRANSAKSI BULANAN - OPERASIONAL & SEWA KENDARAAN (OKTOBER 2026)"
$banner.Font.Bold = $true
$banner.Font.Size = 13
$banner.Font.Color = 0xFFFFFF
$banner.Interior.Color = 0x5C2E0B # Slate / Deep Navy
$ws1.Rows.Item(1).RowHeight = 28
$banner.VerticalAlignment = -4108 # xlCenter
$banner.HorizontalAlignment = -4108

# Headers on Row 2
$headers = @(
    "No Transaksi",
    "Tanggal Transaksi",
    "Customer / Klien",
    "Kategori Layanan",
    "Jenis Unit",
    "Driver / PIC",
    "Wilayah Operasional",
    "Jumlah Unit",
    "Tarif Satuan (Rp)",
    "Total Biaya (Rp)",
    "Status Pembayaran",
    "Catatan"
)

$colLetters = @("A","B","C","D","E","F","G","H","I","J","K","L")
for ($i = 0; $i -lt $headers.Count; $i++) {
    $col = $colLetters[$i]
    $ws1.Range("${col}2").Formula = [string]$headers[$i]
}
$headerRange = $ws1.Range("A2:L2")
$headerRange.Font.Bold = $true
$headerRange.Font.Color = 0xFFFFFF
$headerRange.Interior.Color = 0x8A3A14 # Medium Corporate Blue
$headerRange.HorizontalAlignment = -4108
$headerRange.VerticalAlignment = -4108
$ws1.Rows.Item(2).RowHeight = 22

# Data rows: 41 records
$records = @(
    @("TRX-202610-001", "2026-10-01 08:30:00", "PT Maju Bersama", "Sewa Kendaraan", "Avanza Veloz", "Budi Santoso", "Jakarta Barat", "2", "450000", "Lunas", "VIP Customer"),
    @("TRX-202610-002", "2026-10-01 14:15:00", "CV Sumber Makmur", "Logistik", "Pickup Gran Max", "Ahmad Fauzi", "Bekasi", "1", "350000", "Lunas", "Pengiriman bahan baku"),
    @("TRX-202610-003", "2026-10-02 09:00:00", "PT Indo Jaya Abadi", "Sewa Kendaraan", "Innova Zenix", "Dedi Setiawan", "Jakarta Selatan", "1", "750000", "Pending", "Urgent meeting direksi"),
    @("TRX-202610-004", "2026-10-02 16:45:00", "Toko Berkah Mandiri", "Pengadaan Sparepart", "Forklift 5T", "", "Tangerang", "1", "1200000", "Lunas", "Sparepart genset gudang"),
    @("TRX-202610-005", "2026-10-03 11:20:00", "PT Global Logistik", "Logistik", "Truk Tronton", "Hendra Gunawan", "Surabaya", "3", "2500000", "Cicilan", "Rute lintas Pantura"),
    @("TRX-202610-006", "2026-10-04 10:10:00", "PT Graha Konstruksi", "Sewa Kendaraan", "Truk Tronton", "Budi Santoso", "Bandung", "2", "2500000", "Lunas", "Proyek jembatan tol"),
    @("TRX-202610-007", "2026-10-05 08:00:00", "CV Karya Mandiri", "Perawatan Berkala", "Avanza Veloz", "", "Jakarta Timur", "1", "500000", "Lunas", "Service rutin 20.000 km"),
    @("TRX-202610-008", "2026-10-05 13:30:00", "PT Mitra Niaga", "Sewa Kendaraan", "Innova Zenix", "Ahmad Fauzi", "Jakarta Pusat", "2", "750000", "Pending", "Tamu luar kota"),
    @("TRX-202610-009", "2026-10-06 09:45:00", "PT Maju Bersama", "Logistik", "Pickup Gran Max", "Dedi Setiawan", "Bogor", "1", "350000", "Lunas", "Kirim dokumen legal"),
    @("TRX-202610-010", "2026-10-07 15:20:00", "Toko Sinar Terang", "Konsultasi", "Avanza Veloz", "Hendra Gunawan", "Tangerang", "1", "600000", "Lunas", "Audit operasional rute"),
    @("TRX-202610-011", "2026-10-08 08:15:00", "PT Indo Jaya Abadi", "Sewa Kendaraan", "Forklift 5T", "", "Bekasi", "2", "1200000", "", "Bongkar muat kontainer"),
    @("TRX-202610-012", "2026-10-09 11:00:00", "PT Bintang Timur", "Logistik", "Truk Tronton", "Budi Santoso", "Surabaya", "1", "2500000", "Lunas", "Kargo tekstil ekspor"),
    @("TRX-202610-013", "2026-10-10 14:40:00", "CV Sumber Makmur", "Perawatan Berkala", "Pickup Gran Max", "Ahmad Fauzi", "Jakarta Barat", "1", "400000", "Lunas", "Ganti oli transmisi"),
    @("TRX-202610-014", "2026-10-11 10:30:00", "PT Mitra Niaga", "Sewa Kendaraan", "Innova Zenix", "Dedi Setiawan", "Jakarta Selatan", "1", "750000", "Pending", "Kunjungan kerja Pemda"),
    @("TRX-202610-015", "2026-10-12 08:30:00", "PT Global Logistik", "Logistik", "Truk Tronton", "Hendra Gunawan", "Semarang", "2", "2400000", "Cicilan", "Rute Jawa Tengah"),
    @("TRX-202610-016", "2026-10-12 17:00:00", "PT Maju Bersama", "Sewa Kendaraan", "Avanza Veloz", "Budi Santoso", "Jakarta Timur", "1", "450000", "Lunas", "Transport staf lembur"),
    @("TRX-202610-017", "2026-10-13 09:15:00", "CV Karya Mandiri", "Pengadaan Sparepart", "Pickup Gran Max", "", "Jakarta Utara", "3", "350000", "Lunas", "Beli ban serep 4 unit"),
    @("TRX-202610-018", "2026-10-14 13:45:00", "PT Graha Konstruksi", "Sewa Kendaraan", "Forklift 5T", "Ahmad Fauzi", "Bandung", "1", "1200000", "Batal", "Dibatalkan pihak klien"),
    @("TRX-202610-019", "2026-10-15 10:00:00", "PT Bintang Timur", "Logistik", "Truk Tronton", "Dedi Setiawan", "Surabaya", "2", "2500000", "Lunas", "Ekspedisi mesin pabrik"),
    @("TRX-202610-020", "2026-10-15 16:30:00", "PT Indo Jaya Abadi", "Konsultasi", "Avanza Veloz", "Hendra Gunawan", "Jakarta Pusat", "1", "600000", "Lunas", "Konsultasi SOP fleet"),
    @("TRX-202610-021", "2026-10-16 08:45:00", "Toko Berkah Mandiri", "Sewa Kendaraan", "Innova Zenix", "Budi Santoso", "Bekasi", "1", "750000", "Pending", "Acara keluarga owner"),
    @("TRX-202610-022", "2026-10-17 11:15:00", "PT Mitra Niaga", "Perawatan Berkala", "Forklift 5T", "", "Tangerang", "1", "800000", "Lunas", "Tune up hidrolik pump"),
    @("TRX-202610-023", "2026-10-18 14:00:00", "CV Sumber Makmur", "Logistik", "Pickup Gran Max", "Ahmad Fauzi", "Bogor", "2", "350000", "Lunas", "Distribusi sayur segar"),
    @("TRX-202610-024", "2026-10-19 09:30:00", "PT Maju Bersama", "Sewa Kendaraan", "Avanza Veloz", "Dedi Setiawan", "Jakarta Barat", "1", "450000", "Lunas", "Operasional harian tim sales"),
    @("TRX-202610-025", "2026-10-20 10:45:00", "PT Global Logistik", "Logistik", "Truk Tronton", "Hendra Gunawan", "Solo", "1", "2300000", "Cicilan", "Pengiriman mebel ekspor"),
    @("TRX-202610-026", "2026-10-21 08:15:00", "PT Indo Jaya Abadi", "Sewa Kendaraan", "Innova Zenix", "Budi Santoso", "Jakarta Selatan", "1", "750000", "Lunas", "Meeting direksi holding"),
    @("TRX-202610-027", "2026-10-22 13:00:00", "CV Karya Mandiri", "Sewa Kendaraan", "Forklift 5T", "", "Jakarta Timur", "2", "1200000", "", "Loading baja ringan pabrik"),
    @("TRX-202610-028", "2026-10-23 15:30:00", "PT Graha Konstruksi", "Logistik", "Pickup Gran Max", "Ahmad Fauzi", "Bandung", "1", "380000", "Pending", "Kirim semen mortar"),
    @("TRX-202610-029", "2026-10-24 11:00:00", "Toko Sinar Terang", "Konsultasi", "Avanza Veloz", "Dedi Setiawan", "Tangerang", "1", "600000", "Lunas", "Review efisiensi BBM armada"),
    @("TRX-202610-030", "2026-10-25 09:00:00", "PT Bintang Timur", "Sewa Kendaraan", "Innova Zenix", "Hendra Gunawan", "Surabaya", "1", "800000", "Lunas", "Perjalanan dinas VIP cabang"),
    @("TRX-202610-031", "2026-10-26 08:30:00", "PT Maju Bersama", "Logistik", "Truk Tronton", "Budi Santoso", "Semarang", "3", "2500000", "Cicilan", "Kargo akhir bulan"),
    @("TRX-202610-032", "2026-10-26 14:15:00", "PT Indo Jaya Abadi", "Perawatan Berkala", "Avanza Veloz", "Ahmad Fauzi", "Jakarta Barat", "1", "450000", "Lunas", "Ganti kampas rem"),
    @("TRX-202610-033", "2026-10-27 10:20:00", "CV Sumber Makmur", "Sewa Kendaraan", "Pickup Gran Max", "", "Bekasi", "2", "350000", "Lunas", "Distribusi sembako pasar"),
    @("TRX-202610-034", "2026-10-28 09:00:00", "PT Mitra Niaga", "Sewa Kendaraan", "Forklift 5T", "Dedi Setiawan", "Jakarta Utara", "1", "1200000", "Pending", "Bongkar peti kemas pelabuhan"),
    @("TRX-202610-035", "2026-10-29 11:30:00", "PT Global Logistik", "Logistik", "Truk Tronton", "Hendra Gunawan", "Surabaya", "2", "2500000", "Lunas", "Ekspedisi cepat elektronika"),
    @("TRX-202610-036", "2026-10-29 16:45:00", "Toko Berkah Mandiri", "Sewa Kendaraan", "Innova Zenix", "Budi Santoso", "Jakarta Pusat", "1", "750000", "Lunas", "Acara pernikahan direktur"),
    @("TRX-202610-037", "2026-10-30 08:15:00", "PT Graha Konstruksi", "Pengadaan Sparepart", "Pickup Gran Max", "", "Bandung", "1", "400000", "Lunas", "Emergency parts genset"),
    @("TRX-202610-038", "2026-10-31 10:00:00", "CV Karya Mandiri", "Sewa Kendaraan", "Avanza Veloz", "Ahmad Fauzi", "Jakarta Selatan", "2", "450000", "Lunas", "Weekend rental instansi"),
    @("TRX-202610-039", "2026-10-31 18:30:00", "PT Bintang Timur", "Logistik", "Truk Tronton", "Dedi Setiawan", "Surabaya", "1", "2600000", "Pending", "Penutupan buku ekspedisi"),
    @("TRX-202611-040", "2026-11-01 09:15:00", "PT Indo Jaya Abadi", "Sewa Kendaraan", "Innova Zenix", "Hendra Gunawan", "Jakarta Barat", "1", "750000", "Lunas", "Awal bulan dinas luar"),
    @("TRX-202611-041", "2026-11-02 11:00:00", "PT Maju Bersama", "Logistik", "Truk Tronton", "Budi Santoso", "Semarang", "2", "2500000", "Lunas", "Kargo batch November")
)

$r = 3
foreach ($rec in $records) {
    $ws1.Range("A$r").Formula = [string]$rec[0]
    $ws1.Range("B$r").Formula = [string]$rec[1]
    $ws1.Range("C$r").Formula = [string]$rec[2]
    $ws1.Range("D$r").Formula = [string]$rec[3]
    $ws1.Range("E$r").Formula = [string]$rec[4]
    $ws1.Range("F$r").Formula = [string]$rec[5]
    $ws1.Range("G$r").Formula = [string]$rec[6]
    $ws1.Range("H$r").Formula = [string]$rec[7]
    $ws1.Range("I$r").Formula = [string]$rec[8]
    $ws1.Range("J$r").Formula = "=H$r*I$r"
    $ws1.Range("K$r").Formula = [string]$rec[9]
    $ws1.Range("L$r").Formula = [string]$rec[10]
    
    if ($r % 2 -eq 0) {
        $ws1.Range("A$r`:L$r").Interior.Color = 0xF7F9FA
    }
    $r++
}

$lastRow = $r - 1

# Formatting
$ws1.Range("B3:B$lastRow").NumberFormat = "yyyy-mm-dd hh:mm"
$ws1.Range("H3:H$lastRow").NumberFormat = "#,##0"
$ws1.Range("I3:I$lastRow").NumberFormat = "Rp #,##0"
$ws1.Range("J3:J$lastRow").NumberFormat = "Rp #,##0"

$ws1.Range("A3:A$lastRow").HorizontalAlignment = -4108
$ws1.Range("B3:B$lastRow").HorizontalAlignment = -4108
$ws1.Range("H3:H$lastRow").HorizontalAlignment = -4108
$ws1.Range("K3:K$lastRow").HorizontalAlignment = -4108

# Gridlines & borders
$ws1.Range("A2:L$lastRow").Borders.LineStyle = 1
$ws1.Range("A2:L$lastRow").Borders.Color = 0xD9D9D9

$ws1.Columns.AutoFit()
for ($c = 1; $c -le 12; $c++) {
    $ws1.Columns.Item($c).ColumnWidth = [Math]::Max($ws1.Columns.Item($c).ColumnWidth + 3, 13)
}

# Sheet 2: Master_Armada
$ws2 = $wb.Worksheets.Add([System.Type]::Missing, $ws1)
$ws2.Name = "Master_Armada"
$ws2.Range("A1").Formula = "ID Unit"
$ws2.Range("B1").Formula = "Model Kendaraan"
$ws2.Range("C1").Formula = "Kapasitas (Ton)"
$ws2.Range("D1").Formula = "Status Unit"

$armada = @(
    @("UNT-01", "Avanza Veloz", "0.5", "Aktif"),
    @("UNT-02", "Innova Zenix", "0.7", "Aktif"),
    @("UNT-03", "Pickup Gran Max", "1.2", "Aktif"),
    @("UNT-04", "Forklift 5T", "5.0", "Perawatan"),
    @("UNT-05", "Truk Tronton", "25.0", "Aktif")
)
$ar = 2
foreach ($row in $armada) {
    $ws2.Range("A$ar").Formula = [string]$row[0]
    $ws2.Range("B$ar").Formula = [string]$row[1]
    $ws2.Range("C$ar").Formula = [string]$row[2]
    $ws2.Range("D$ar").Formula = [string]$row[3]
    $ar++
}
$ws2.Range("A1:D1").Font.Bold = $true
$ws2.Range("A1:D1").Interior.Color = 0x8A3A14
$ws2.Range("A1:D1").Font.Color = 0xFFFFFF
$ws2.Columns.AutoFit()

# Activate Sheet 1
$ws1.Activate()

# Save
$samplePath = (Join-Path (Get-Location) "Sample_Data_Bulanan.xlsx")
$wb.SaveAs($samplePath, 51)
$wb.Close($false)
$xl.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
Write-Host "SUCCESS: Sample data created at: $samplePath"
