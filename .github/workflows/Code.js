function generatePDFDCAFloq() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Input DCA");
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert("Tab 'Input DCA' tidak ditemukan!");
    return;
  }
  
  var data = sheet.getDataRange().getValues();
  var totalBtc = 0;
  var totalModal = 0;
  var rowsHtml = "";
  
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;

    var tanggal = data[i][0] instanceof Date ? 
        Utilities.formatDate(data[i][0], "GMT+7", "dd/MM/yyyy") : data[i][0];
    
    var namaAset = data[i][1] ? data[i][1] : "BTC";
    var jenis = data[i][2] ? data[i][2] : "BUY";
    var jumlahOrderNum = Number(data[i][3]) || 0;
    var hargaBeliNum   = Number(data[i][4]) || 0;
    var btcDiterimaNum = Number(data[i][7]) || 0;

    totalBtc += btcDiterimaNum;
    totalModal += jumlahOrderNum;

    var jumlahOrderStr = "Rp" + jumlahOrderNum.toLocaleString("id-ID");
    var hargaBeliStr   = "Rp" + hargaBeliNum.toLocaleString("id-ID");
    var btcDiterimaStr = btcDiterimaNum.toFixed(8) + " BTC";

    rowsHtml += `
      <tr>
        <td style="text-align: center;">${tanggal}</td>
        <td style="text-align: center;">${namaAset}</td>
        <td style="text-align: center;">${jenis}</td>
        <td style="text-align: right;">${jumlahOrderStr}</td>
        <td style="text-align: right;">${hargaBeliStr}</td>
        <td style="text-align: right;">${btcDiterimaStr}</td>
      </tr>`;
  }

  var avgHargaBeli = totalBtc > 0 ? (totalModal / totalBtc) : 0;

  var htmlContent = `
  <!DOCTYPE html>
  <html>
  <head>
    <style>
      @page { size: A4 portrait; margin: 0; }
      body { font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif; margin: 0; padding: 0; color: #333; }
      .header-card { width: 100%; border-bottom: 2px solid #0d3838; padding: 25px 30px; box-sizing: border-box; }
      .header-table { width: 100%; border-collapse: collapse; }
      .header-table td { vertical-align: top; }
      .brand-section { width: 50%; }
      .floq-logo-text { font-size: 24pt; font-weight: 800; color: #000; letter-spacing: 2px; }
      .company-subtitle { font-size: 9pt; font-weight: 600; color: #555; margin-top: 5px; }
      .doc-badge { display: inline-block; margin-top: 12px; background-color: #e8f4f1; color: #0d3838; font-size: 8.5pt; font-weight: 700; padding: 4px 10px; border-radius: 4px; border-left: 3px solid #1a4d40; }
      .owner-section { width: 50%; padding-left: 15px; }
      .owner-card-inner { background-color: #f9fbfb; border: 1px solid #e1ecea; border-radius: 6px; padding: 12px 16px; }
      .owner-title { font-size: 10pt; font-weight: 700; color: #0d3838; margin-bottom: 8px; border-bottom: 1px solid #e1ecea; padding-bottom: 4px; }
      .info-grid { width: 100%; border-collapse: collapse; }
      .info-grid td { padding: 3px 0; font-size: 8.5pt; vertical-align: top; }
      .lbl { color: #666; width: 75px; }
      .cln { width: 10px; color: #666; text-align: center; }
      .val { font-weight: 700; color: #111; }
      .content { padding: 25px 30px; }
      .section-title { font-size: 12pt; font-weight: 700; color: #111; margin-bottom: 10px; margin-top: 10px; }
      .data-table { width: 100%; border-collapse: collapse; margin-top: 8px; margin-bottom: 20px; }
      .data-table th { background-color: #0d3838; color: #ffffff; font-size: 8.5pt; font-weight: 700; padding: 8px 10px; }
      .data-table td { font-size: 8.5pt; padding: 7px 10px; border-bottom: 1px solid #eeeeee; }
    </style>
  </head>
  <body>
    <div class="header-card">
      <table class="header-table">
        <tr>
          <td class="brand-section">
            <div class="floq-logo-text">FLOQ</div>
            <div class="company-subtitle">Floq (PT Kripto Maksima Koin)</div>
            <div class="doc-badge">LAPORAN AKUMULASI DCA BITCOIN</div>
          </td>
          <td class="owner-section">
            <div class="owner-card-inner">
              <div class="owner-title">Informasi Pemilik</div>
              <table class="info-grid">
                <tr><td class="lbl">Nama</td><td class="cln">:</td><td class="val">NURIL AMANDA</td></tr>
                <tr><td class="lbl">KTP / NPWP</td><td class="cln">:</td><td class="val">1101031201040001</td></tr>
                <tr><td class="lbl">Status</td><td class="cln">:</td><td class="val" style="color: #0f766e;">Active Accumulator</td></tr>
                <tr><td class="lbl">Periode</td><td class="cln">:</td><td class="val">Juni 2025 – Desember 2026</td></tr>
              </table>
            </div>
          </td>
        </tr>
      </table>
    </div>

    <div class="content">
      <div class="section-title">Ringkasan Portofolio</div>
      <table class="data-table">
        <thead>
          <tr>
            <th style="text-align: left;">Aset</th>
            <th style="text-align: right;">Total Akumulasi</th>
            <th style="text-align: right;">Harga Rata-Rata Beli</th>
            <th style="text-align: right;">Total Modal (IDR)</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td style="text-align: left;">Bitcoin (BTC)</td>
            <td style="text-align: right;">${totalBtc.toFixed(8)} BTC</td>
            <td style="text-align: right;">Rp${Math.round(avgHargaBeli).toLocaleString("id-ID")}</td>
            <td style="text-align: right;">Rp${totalModal.toLocaleString("id-ID")}</td>
          </tr>
        </tbody>
      </table>

      <div class="section-title">Rincian Transaksi DCA (Data Google Sheets)</div>
      <table class="data-table">
        <thead>
          <tr>
            <th style="text-align: center;">Tanggal</th>
            <th style="text-align: center;">Nama Aset</th>
            <th style="text-align: center;">Jenis</th>
            <th style="text-align: right;">Jumlah Order</th>
            <th style="text-align: right;">Harga Beli</th>
            <th style="text-align: right;">BTC Diterima</th>
          </tr>
        </thead>
        <tbody>
          ${rowsHtml ? rowsHtml : '<tr><td colspan="6" style="text-align:center;">Belum ada data transaksi.</td></tr>'}
        </tbody>
      </table>
    </div>
  </body>
  </html>`;

  var blob = Utilities.newBlob(htmlContent, "text/html", "Laporan_DCA_Floq.html");
  var pdf = blob.getAs("application/pdf").setName("Laporan_DCA_Floq_Personal.pdf");
  
  var file = DriveApp.createFile(pdf);
  SpreadsheetApp.getUi().alert("Sukses! PDF Laporan DCA tersimpan di Google Drive dengan nama: " + file.getName());
}
  
