// ==========================================
// 1. SETUP & UTILITIES
// ==========================================

function doGet() {
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('SiGAS PRO - Sistem Agen Gas')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = [
    {name: 'USERS', header: ['Username', 'Password', 'Role', 'Nama']},
    {name: 'PRODUK', header: ['ID', 'Nama_Produk', 'Harga_Jual', 'Harga_Beli', 'Stok_Isi', 'Stok_Kosong', 'SKU', 'Kode', 'Link_Gambar', 'Harga_Beli_Tabung', 'Harga_Jual_Tabung', 'Harga_Jual_Refill_2', 'Harga_Jual_Refill_3', 'Harga_Jual_Tabung_2', 'Harga_Jual_Tabung_3']},
    {name: 'PELANGGAN', header: ['ID', 'Nama', 'Nama_Perusahan', 'NoHP', 'Alamat']},
    {name: 'SUPPLIER', header: ['ID', 'Nama_Supplier', 'NoHP', 'Alamat']},
    {name: 'TRANSAKSI', header: ['ID_Trans', 'Waktu', 'Pelanggan', 'Produk', 'Qty', 'Total', 'Tipe', 'Kasir', 'Metode_Bayar', 'Jatuh_Tempo', 'Status', 'Qty_Retur']},
    {name: 'PEMBELIAN', header: ['ID_Beli', 'Waktu', 'Supplier', 'Produk', 'Qty', 'Total', 'Metode']},
    {name: 'KEUANGAN', header: ['ID', 'Tanggal', 'Jenis', 'Kategori', 'Nominal', 'Keterangan', 'Akun']}, 
    {name: 'KATEGORI', header: ['Nama_Kategori']},
    {name: 'KARYAWAN', header: ['ID', 'Nama_Lengkap', 'Tempat_Lahir', 'Tanggal_Lahir', 'Jenis_Kelamin', 'No_Identitas', 'Tipe_Identitas', 'Email', 'Alamat_KTP', 'Alamat_Domisili', 'Nama_Darurat', 'Telp_Darurat', 'Gaji_Pokok', 'Bonus', 'Foto_Url', 'Status', 'Tanggal_Masuk']}, 
    {name: 'RIWAYAT_GAJI', header: ['ID_Gaji', 'Periode', 'Tanggal_Generate', 'Nama_Karyawan', 'Gaji_Pokok', 'Bonus', 'Potongan_Kasbon', 'Total_THP', 'Status', 'Kasbon_Baru']}, 
    {name: 'KASBON', header: ['ID_Kasbon', 'Tanggal', 'Nama_Karyawan', 'Nominal', 'Keterangan', 'Status_Lunas', 'Sudah_Bayar', 'Tenor', 'Angsuran_Per_Bulan']},
    {name: 'RIWAYAT_BAYAR_KASBON', header: ['ID_Bayar', 'ID_Kasbon', 'Tanggal_Bayar', 'Nominal_Bayar', 'Tipe_Bayar', 'Keterangan']},
    {name: 'PENGATURAN', header: ['Key', 'Value']},
    {name: 'AKUN_KAS', header: ['ID_Akun', 'Nama_Akun', 'No_Rekening', 'Tipe', 'Saldo_Awal']},
    {name: 'RIWAYAT_SO', header: ['ID_SO', 'TANGGAL', 'PETUGAS', 'TOTAL_ITEM', 'DATA_JSON']},
    {name: 'LOG_AKTIVITAS', header: ['TIMESTAMP', 'USER / EMAIL', 'JENIS AKSI', 'DETAIL AKTIVITAS']} 
  ];

  sheets.forEach(s => {
    let sheet = ss.getSheetByName(s.name);
    if (!sheet) {
      sheet = ss.insertSheet(s.name);
      sheet.appendRow(s.header);
      
      // Data Default User
      if(s.name === 'USERS') sheet.appendRow(['admin', 'admin123', 'Admin', 'Super Admin']);
      
      // Data Default Akun Kas
      if(s.name === 'AKUN_KAS') {
         sheet.appendRow(['ACC-1', 'Kas Tunai (Laci)', '-', 'Tunai', 0]); 
         sheet.appendRow(['ACC-2', 'Bank BCA', '', 'Bank', 0]);
         sheet.appendRow(['ACC-3', 'Bank BRI', '', 'Bank', 0]);
         sheet.appendRow(['ACC-4', 'Bank BSI', '', 'Bank', 0]);
         sheet.appendRow(['ACC-5', 'Bank MANDIRI', '', 'Bank', 0]);
      }
    }
  });
}

// [FIX TOTAL] Fungsi Baca Data yang Aman dari Error Timezone
function getData(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  
  // 1. Cek Apakah Sheet Ada?
  if (!sheet) return [];

  // 2. Ambil Range Data
  const range = sheet.getDataRange();
  const values = range.getValues();
  
  // 3. Cek Apakah Data Kosong?
  if (!values || values.length <= 1) return [];

  // 4. Ambil Timezone Spreadsheet
  const timeZone = ss.getSpreadsheetTimeZone();

  // 5. Filter baris kosong & Format Tanggal
  const data = values.slice(1);
  return data.filter(r => r[0] !== "").map(r => {
      return r.map(cell => {
          if (cell instanceof Date) {
              return Utilities.formatDate(cell, timeZone, "yyyy-MM-dd'T'HH:mm:ss");
          }
          return cell;
      });
  });
}

function getDaftarAkun() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetAkun = ss.getSheetByName('AKUN_KAS');
  const sheetKeu = ss.getSheetByName('KEUANGAN');
   
  if(!sheetAkun || !sheetKeu) return [];

  const dataAkun = sheetAkun.getDataRange().getValues().slice(1);
  const dataKeu = sheetKeu.getDataRange().getValues().slice(1);

  let listAkun = dataAkun.map(r => {
      let id = r[0];
      let nama = r[1];
      let norek = r[2]; 
      let tipe = r[3];  
      let saldo = Number(r[4]); 

      // Loop Transaksi Keuangan untuk hitung saldo realtime
      dataKeu.forEach(k => {
          let akunTrx = k[6]; 
          let jenis = k[2];
          let nominal = Number(k[4]);

          if(akunTrx === nama) {
              if(jenis === 'Pemasukan') saldo += nominal;
              if(jenis === 'Pengeluaran') saldo -= nominal;
          }
      });

      return { id: id, nama: nama, norek: norek, tipe: tipe, saldo: saldo };
  });

  return listAkun;
}

// ==========================================
// 2. USER & AUTH
// ==========================================

function loginUser(username, password) {
  const data = getData('USERS');
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] == username && data[i][1] == password) {
      return { status: 'success', role: data[i][2], nama: data[i][3], username: data[i][0] }; 
    }
  }
  return { status: 'failed' };
}

function getAllUsers() {
  return getData('USERS');
}

function simpanUserBaru(form) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('USERS');
  const data = sheet.getDataRange().getValues();

  let userExists = false;
  let rowIndex = -1;

  for(let i=1; i<data.length; i++) {
     if(data[i][0] === form.username) {
        userExists = true;
        rowIndex = i + 1;
        break;
     }
  }

  if(form.isEdit && userExists) {
     let oldPass = sheet.getRange(rowIndex, 2).getValue();
     let newPass = form.password ? form.password : oldPass;
     
     sheet.getRange(rowIndex, 1, 1, 4).setValues([[form.username, newPass, form.role, form.nama]]);
     return "Data User Berhasil Diupdate";
  } else if (!form.isEdit && userExists) {
     return "Error: Username sudah terpakai!";
  } else {
     sheet.appendRow([form.username, form.password, form.role, form.nama]);
     return "User Baru Berhasil Ditambahkan";
  }
}

function hapusUser(username) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('USERS');
  const data = sheet.getDataRange().getValues();
  for(let i=1; i<data.length; i++) {
    if(data[i][0] == username) {
       sheet.deleteRow(i+1);
       return "User dihapus.";
    }
  }
}

function gantiPasswordSendiri(username, oldPass, newPass) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('USERS');
  const data = sheet.getDataRange().getValues();
   
  for(let i=1; i<data.length; i++) {
    if(data[i][0] == username) {
       if(data[i][1] != oldPass) return "Password Lama Salah!";
       
       sheet.getRange(i+1, 2).setValue(newPass);
       return "Password Berhasil Diganti";
    }
  }
  return "User tidak ditemukan";
}

// ==========================================
// 3. SETTINGS & PROFIL
// ==========================================

function getProfilPerusahaan() {
  const data = getData('PENGATURAN');
  let config = {};
  data.forEach(row => {
     config[row[0]] = row[1];
  });
  return config;
}

function simpanProfilPerusahaan(form) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('PENGATURAN');
  const data = sheet.getDataRange().getValues();
   
  const FOLDER_ID = '15hiLtvusofF2OJpXVq8lJkePbmqVIuPM'; 

    const updateOrInsert = (key, val) => {
     let found = false;
     
     let finalVal = val;
     if (key === 'no_perusahaan' || key === 'no_pemilik') {
         finalVal = "'" + val; 
     }

     for(let i=1; i<data.length; i++) {
        if(data[i][0] === key) {
           sheet.getRange(i+1, 2).setValue(finalVal);
           found = true;
           break;
        }
     }
     if(!found) sheet.appendRow([key, finalVal]);
  };

  // 1. PROSES UPLOAD LOGO
  if (form.logo && form.logo.data) {
    try {
       const decoded = Utilities.base64Decode(form.logo.data);
       const blob = Utilities.newBlob(decoded, form.logo.mimeType, 'LOGO-' + Date.now());
       
       const folder = DriveApp.getFolderById(FOLDER_ID);
       const file = folder.createFile(blob);
       
       file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
       const logoUrl = "https://drive.google.com/thumbnail?id=" + file.getId() + "&sz=w1000";
       
       updateOrInsert('logo_perusahaan', logoUrl);

    } catch (e) {
       throw new Error("Gagal Upload Logo: " + e.message);
    }
  }

  // 2. Simpan Data Teks
  updateOrInsert('nama_perusahaan', form.nama_perusahaan);
  updateOrInsert('nama_pemilik', form.nama_pemilik);
  updateOrInsert('alamat', form.alamat);
  updateOrInsert('no_perusahaan', form.no_perusahaan);
  updateOrInsert('no_pemilik', form.no_pemilik);

  catatLog("SETTING", "Mengubah profil perusahaan");

  return "Profil & Logo Berhasil Disimpan!";
}

// ==========================================
// 4. PRODUK
// ==========================================

function tambahProduk(form) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('PRODUK');
   
  const FOLDER_ID = '15hiLtvusofF2OJpXVq8lJkePbmqVIuPM'; 
   
  let imageUrl = '';

  // PROSES UPLOAD
  if (form.gambar && form.gambar.data) {
    try {
      const decoded = Utilities.base64Decode(form.gambar.data);
      const blob = Utilities.newBlob(decoded, form.gambar.mimeType, form.gambar.fileName);
      
      const folder = DriveApp.getFolderById(FOLDER_ID);
      const file = folder.createFile(blob); 
      
      try {
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      } catch (e1) {
        // Fallback jika gagal public
        try {
           file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
        } catch (e2) {
           console.log("Gagal set permission: " + e1.message); 
        }
      }

      imageUrl = "https://drive.google.com/thumbnail?id=" + file.getId() + "&sz=w1000";

    } catch (e) {
      throw new Error("Gagal Upload: " + e.message); 
    }
  } else {
    imageUrl = (typeof form.gambar === 'string') ? form.gambar : '';
  }

  // Simpan data
// Simpan data (UPDATE KOLOM BARU)
  sheet.appendRow([
    'P-' + Date.now(), 
    form.nama, 
    form.hargaJual,       
    form.hargaBeli,       
    form.stokIsi, 
    form.stokKosong,
    form.sku,      
    form.kode,     
    imageUrl,
    form.hargaBeliTabung || 0, 
    form.hargaJualTabung || 0,
    // --- 4 KOLOM BARU ---
    form.hargaJualRefill2 || 0, // Index 11
    form.hargaJualRefill3 || 0, // Index 12
    form.hargaJualTabung2 || 0, // Index 13
    form.hargaJualTabung3 || 0  // Index 14
  ]);

  // [FIX] formObject -> form
  catatLog("PRODUK_BARU", "Menambah produk: " + form.nama + " | Stok Awal: " + form.stokIsi);
}

function updateProduk(form) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('PRODUK');
  const data = sheet.getDataRange().getValues();
   
  const FOLDER_ID = '15hiLtvusofF2OJpXVq8lJkePbmqVIuPM'; 

  let rowTarget = -1;
  let oldImage = '';

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == form.id) {
      rowTarget = i + 1;
      oldImage = data[i][8]; 
      break;
    }
  }

  if (rowTarget === -1) throw new Error("Produk tidak ditemukan/ID salah.");

  let finalImageUrl = oldImage;

  if (form.gambar && form.gambar.data) {
    try {
      const decoded = Utilities.base64Decode(form.gambar.data);
      const blob = Utilities.newBlob(decoded, form.gambar.mimeType, 'UPD-' + form.gambar.fileName);
      const folder = DriveApp.getFolderById(FOLDER_ID);
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      // [FIX] finalimageUrl -> finalImageUrl
      finalImageUrl = "https://drive.google.com/uc?export=view&id=" + file.getId();
    } catch (e) {
      console.log("Gagal update gambar: " + e.message);
    }
  } else if (typeof form.gambar === 'string' && form.gambar !== '') {
      finalImageUrl = form.gambar;
  }

  sheet.getRange(rowTarget, 2).setValue(form.nama);
  sheet.getRange(rowTarget, 3).setValue(form.hargaJual);
  sheet.getRange(rowTarget, 4).setValue(form.hargaBeli);
  sheet.getRange(rowTarget, 7).setValue(form.sku);
  sheet.getRange(rowTarget, 8).setValue(form.kode);
  sheet.getRange(rowTarget, 9).setValue(finalImageUrl);
   
  sheet.getRange(rowTarget, 10).setValue(form.hargaBeliTabung || 0); 
  sheet.getRange(rowTarget, 11).setValue(form.hargaJualTabung || 0);

  // --- UPDATE KOLOM BARU ---
  sheet.getRange(rowTarget, 12).setValue(form.hargaJualRefill2 || 0);
  sheet.getRange(rowTarget, 13).setValue(form.hargaJualRefill3 || 0);
  sheet.getRange(rowTarget, 14).setValue(form.hargaJualTabung2 || 0);
  sheet.getRange(rowTarget, 15).setValue(form.hargaJualTabung3 || 0);

  // [FIX] formObject -> form
  catatLog("EDIT_PRODUK", "Mengupdate data produk: " + form.nama);

  return "Produk Berhasil Diupdate!";
}

function hapusProduk(nama) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PRODUK');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] == nama) { sheet.deleteRow(i + 1); break; }
  }
}

// ==========================================
// 5. TRANSAKSI (PENJUALAN)
// ==========================================

function simpanTransaksiBulk(dataTransaksi) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const prodSheet = ss.getSheetByName('PRODUK');
  const trxSheet = ss.getSheetByName('TRANSAKSI');
  const keuSheet = ss.getSheetByName('KEUANGAN');
  
  // Pastikan Header punya kolom Diskon
if (trxSheet.getLastColumn() < 12) {
    trxSheet.getRange(1, 12).setValue('Diskon');
}
// >>> TAMBAHKAN INI <<<
if (trxSheet.getLastColumn() < 13) {
    trxSheet.getRange(1, 13).setValue('Qty_Retur');
}
   
  const prodData = prodSheet.getDataRange().getValues();
  const idTrxMaster = 'KBA-' + Date.now();
  const waktu = new Date();
   
  let totalBelanja = 0;
  let summaryProduk = [];
  let updatesToExecute = [];

  // --- TAHAP 1: VALIDASI STOK (Sama seperti sebelumnya) ---
  for (let x = 0; x < dataTransaksi.items.length; x++) {
      let item = dataTransaksi.items[x];
      let itemFound = false;

      for (let i = 1; i < prodData.length; i++) {
          if (prodData[i][1] == item.produkNama) {
              let curIsi = Number(prodData[i][4]);    
              let curKosong = Number(prodData[i][5]); 
              let qty = Number(item.qty);
              let tipe = String(item.tipe).toLowerCase();

              let newIsi = curIsi;
              let newKosong = curKosong;

              if (tipe.includes('kosong')) {
                  if (curKosong < qty) throw new Error(`Stok Wadah ${item.produkNama} Kurang! Sisa: ${curKosong}`);
                  newKosong = curKosong - qty;
              } else if (tipe.includes('baru') || (tipe.includes('isi') && tipe.includes('wadah'))) {
                  if (curIsi < qty) throw new Error(`Stok Isi ${item.produkNama} Kurang! Sisa: ${curIsi}`);
                  newIsi = curIsi - qty;
              } else {
                  if (curIsi < qty) throw new Error(`Stok Isi ${item.produkNama} Kurang! Sisa: ${curIsi}`);
                  newIsi = curIsi - qty;
                  newKosong = curKosong + qty; 
              }

              updatesToExecute.push({
                  rowIndex: i + 1,
                  valIsi: newIsi,
                  valKosong: newKosong,
                  nama: item.produkNama
              });
              
              prodData[i][4] = newIsi;
              prodData[i][5] = newKosong;
              itemFound = true;
              break;
          }
      }
      if(!itemFound) throw new Error(`Produk ${item.produkNama} tidak ditemukan di Database.`);
  }

  // --- TAHAP 2: EKSEKUSI STOK (Sama seperti sebelumnya) ---
  updatesToExecute.forEach(upd => {
      prodSheet.getRange(upd.rowIndex, 5).setValue(upd.valIsi);    
      prodSheet.getRange(upd.rowIndex, 6).setValue(upd.valKosong); 
  });

  // --- TAHAP 3: CATAT TRANSAKSI (UPDATE DISINI) ---
  let statusTrx = (dataTransaksi.metode === 'Hutang') ? 'Belum Lunas' : 'Lunas';
  
  // Ambil nilai diskon (jika ada)
  let nilaiDiskon = Number(dataTransaksi.diskon) || 0;
  
  // Kita simpan diskon hanya di baris item PERTAMA agar tidak terhitung ganda saat dijumlah
  let isFirstItem = true;

  dataTransaksi.items.forEach(item => {
      let diskonBaris = isFirstItem ? nilaiDiskon : 0; // Tempel diskon di baris pertama saja
      
      trxSheet.appendRow([
        idTrxMaster, 
        waktu, 
        dataTransaksi.pelanggan, 
        item.produkNama, 
        item.qty, 
        item.total, 
        item.tipe, 
        dataTransaksi.kasir, 
        dataTransaksi.metode, 
        dataTransaksi.jatuhTempo, 
        statusTrx,
        diskonBaris // <--- [BARU] Simpan Diskon di Kolom 12 (Index 11)
      ]);
      
      totalBelanja += Number(item.total);
      summaryProduk.push(`${item.produkNama} (${item.qty})`);
      isFirstItem = false;
  });

  // Catat Keuangan
  if (dataTransaksi.metode !== 'Hutang') {
    let ket = `Penjualan Gas`;
    if(nilaiDiskon > 0) {
        ket += ` (Disc: ${nilaiDiskon})`;
    }
    // Gunakan totalFinal jika ada, atau hitung manual
    let nominalMasuk = (dataTransaksi.totalFinal !== undefined) ? dataTransaksi.totalFinal : (totalBelanja - nilaiDiskon);

    keuSheet.appendRow([
        'FIN-' + idTrxMaster, waktu, 'Pemasukan', 'Penjualan Gas', 
        nominalMasuk, `Penjualan: ${summaryProduk.join(', ')}. ${ket}`, dataTransaksi.metode 
    ]);
  }

  catatLog("TRANSAKSI_BARU", "ID: " + idTrxMaster + " | Pelanggan: " + dataTransaksi.pelanggan + " | Total Bersih: " + (totalBelanja - nilaiDiskon));
   
  SpreadsheetApp.flush(); 
  return "Transaksi Berhasil Disimpan!";
}

function getDataPiutang() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TRANSAKSI');
  if (!sheet) return [];
  const allData = sheet.getDataRange().getValues();
  if (allData.length < 2) return [];

  const idxMetode = 8;
  const idxJatuhTempo = 9;
  const idxStatus = 10;

  let grouped = {};

  for (let i = 1; i < allData.length; i++) {
    let row = allData[i];
    let metode = String(row[idxMetode]).trim();
    
    if (metode === 'Hutang') {
       let id = row[0];
       let status = String(row[idxStatus]).trim(); 

       if(!grouped[id]) {
          let tglWaktu = (row[1] instanceof Date) ? row[1].toISOString() : String(row[1]);
          let tglTempo = (row[idxJatuhTempo] instanceof Date) ? row[idxJatuhTempo].toISOString() : String(row[idxJatuhTempo]);
          
          grouped[id] = {
             id: id,
             waktu: tglWaktu,       
             pelanggan: row[2],
             total: 0,
             jatuhTempo: tglTempo,
             status: status 
          };
       }
       grouped[id].total += Number(row[5]);
    }
  }
   
  return Object.values(grouped).sort((a, b) => {
      if (a.status === b.status) {
          return new Date(b.waktu) - new Date(a.waktu); 
      }
      return a.status === 'Belum Lunas' ? -1 : 1;
  }).map(x => [x.id, x.waktu, x.pelanggan, x.total, x.jatuhTempo, x.status]);
}

function lunasiHutang(idTrx, totalBayar, namaPelanggan) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetTrx = ss.getSheetByName('TRANSAKSI');
  const sheetKeu = ss.getSheetByName('KEUANGAN');
   
  const dataTrx = sheetTrx.getDataRange().getValues();
   
  for(let i=1; i<dataTrx.length; i++) {
     if(dataTrx[i][0] == idTrx) {
        sheetTrx.getRange(i+1, 11).setValue('Lunas'); 
     }
  }

  sheetKeu.appendRow([
      'LUNAS-' + Date.now(), 
      new Date(), 
      'Pemasukan', 
      'Pelunasan Piutang', 
      totalBayar, 
      `Pelunasan Bon: ${namaPelanggan} (${idTrx})`
  ]);

  return "Hutang Berhasil Dilunasi & Masuk Kas!";
}

function getJumlahJatuhTempo() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TRANSAKSI');
  const data = sheet.getDataRange().getValues();
  const today = new Date();
  let count = 0;
  let uniqueIDs = []; 

  for (let i = 1; i < data.length; i++) {
    let idTrx = data[i][0];
    let status = data[i][10]; 
    let tglTempo = new Date(data[i][9]); 

    if (status === 'Belum Lunas' && tglTempo <= today && !uniqueIDs.includes(idTrx)) {
       count++;
       uniqueIDs.push(idTrx);
    }
  }
  return count;
}

function getRiwayatTransaksi() {
  const data = getData('TRANSAKSI');
   
  let grouped = {};
  data.forEach(row => {
    let id = row[0];
    let waktuStr = row[1] instanceof Date ? row[1].toISOString() : row[1];

    if (!grouped[id]) {
      grouped[id] = {
        id: id,
        waktu: waktuStr,
        pelanggan: row[2],
        kasir: row[7],
        metode: row[8],        
        jatuhTempo: row[9],  
        totalBayar: 0,  
        items: [],
        diskonTotal: 0 
      };
    }
    
    grouped[id].items.push({
      produk: row[3],
      qty: row[4],
      hargaTotal: row[5],
      tipe: row[6],
      status: row[10], // Index 10 = Status
      
      // >>> TAMBAHKAN BARIS DI BAWAH INI <<<
      // Baca kolom 13 (Index 12) supaya HTML tahu ada retur 1 pcs
      qtyRetur: row[12] 
    });

    grouped[id].totalBayar += Number(row[5]);
    
    if(row[11]) {
        grouped[id].diskonTotal += Number(row[11]);
    }
  });

  let result = Object.values(grouped).map(trx => {
      trx.totalBayar = trx.totalBayar - trx.diskonTotal;
      
      // [BARU] Cek apakah ada item yang statusnya 'Retur'
      trx.isRetur = trx.items.some(item => String(item.status).toLowerCase().includes('retur'));
      
      return trx;
  });

  return result.sort((a, b) => new Date(b.waktu) - new Date(a.waktu)).slice(0, 50);
}

// ==========================================
// 6. RETUR & HAPUS
// ==========================================

// [CODE.GS] Cari fungsi ini dan ganti isinya
function prosesReturBaru(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const prodSheet = ss.getSheetByName('PRODUK');
  const keuSheet = ss.getSheetByName('KEUANGAN');
  const trxSheet = ss.getSheetByName('TRANSAKSI'); // Tambahan: Load sheet Transaksi
   
  const prodData = prodSheet.getDataRange().getValues();
  // Tambahan: Ambil data transaksi untuk pencarian baris
  const trxData = trxSheet.getDataRange().getValues();

  let totalRefund = 0;
  let totalHPPRetur = 0;
  
  // Loop item yang diretur
  payload.items.forEach(returItem => {
    if(returItem.qtyRetur > 0) {
      
      // 1. UPDATE STOK PRODUK
      for (let i = 1; i < prodData.length; i++) {
        if (prodData[i][1] == returItem.produk) {
           let curIsi = Number(prodData[i][4]);
           let curKosong = Number(prodData[i][5]);
           let hppIsi = Number(prodData[i][3]) || 0;
           let hppTabung = Number(prodData[i][9]) || 0;
           
           if(payload.jenis === 'JUAL') {
              // Balikin Stok
              prodSheet.getRange(i+1, 5).setValue(curIsi + returItem.qtyRetur);
              
              if(returItem.tipe && returItem.tipe.includes('Refill')) {
                 prodSheet.getRange(i+1, 6).setValue(curKosong - returItem.qtyRetur);
                 totalHPPRetur += (hppIsi * returItem.qtyRetur);
              } else {
                 if(returItem.tipe && returItem.tipe.toLowerCase().includes('baru')) {
                     totalHPPRetur += ((hppIsi + hppTabung) * returItem.qtyRetur);
                 } else {
                     totalHPPRetur += (hppIsi * returItem.qtyRetur);
                 }
              }

              // --- [BARU] UPDATE STATUS DI TRANSAKSI JADI 'Retur' ---
                for(let r = 1; r < trxData.length; r++) {
                    // ... di dalam loop for ...
                  if(trxData[r][0] == payload.idTrx && trxData[r][3] == returItem.produk) {
                      
                      // Update Status
                      trxSheet.getRange(r+1, 11).setValue('Retur');
                      
                      // >>> TAMBAHKAN INI (Simpan Angka Retur) <<<
                      // Simpan di Kolom 13 (M)
                      trxSheet.getRange(r+1, 13).setValue(returItem.qtyRetur); 
                      
                      break; 
                  }
                }
              // ------------------------------------------------------

           } else {
              // Retur Pembelian
              prodSheet.getRange(i+1, 5).setValue(curIsi - returItem.qtyRetur);
           }
           break;
        }
      }

      totalRefund += (returItem.hargaSatuan * returItem.qtyRetur);
    }
  });

  // Catat ke Keuangan
  if(totalRefund > 0) {
     if(payload.jenis === 'JUAL') {
        keuSheet.appendRow([
            'RET-' + Date.now(), 
            new Date(), 
            'Pengeluaran', 
            'Retur Penjualan', 
            totalRefund, 
            `Retur TRX: ${payload.idTrx}. ${payload.alasan}`,
            totalHPPRetur 
        ]);
     } else {
        keuSheet.appendRow([
            'RET-' + Date.now(), 
            new Date(), 
            'Pemasukan', 
            'Retur Pembelian', 
            totalRefund, 
            `Retur BELI: ${payload.idTrx}. ${payload.alasan}`
        ]);
     }
  }

  SpreadsheetApp.flush(); 
  return "Retur Berhasil Diproses!";
}

function prosesRetur(idTrx, produkNama, qty, tipe, mode) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const prodSheet = ss.getSheetByName('PRODUK');
  const trxSheet = ss.getSheetByName('TRANSAKSI');
  const keuSheet = ss.getSheetByName('KEUANGAN');
   
  const prodData = prodSheet.getDataRange().getValues();
  for (let i = 1; i < prodData.length; i++) {
    if (prodData[i][1] == produkNama) {
       let curIsi = Number(prodData[i][4]);
       let curKosong = Number(prodData[i][5]);
       
       prodSheet.getRange(i + 1, 5).setValue(curIsi + Number(qty));
       
       if(tipe === 'Tukar (Refill)') {
          prodSheet.getRange(i + 1, 6).setValue(curKosong - Number(qty));
       }
       break;
    }
  }

  const trxData = trxSheet.getDataRange().getValues();
  let nominalRefund = 0;

  for(let i=1; i<trxData.length; i++) {
    if(trxData[i][0] == idTrx && trxData[i][3] == produkNama && trxData[i][8] != 'Retur') {
       if(mode === 'FULL') {
         trxSheet.deleteRow(i+1); 
       } else {
         trxSheet.getRange(i+1, 9).setValue('Retur Item');
       }
       nominalRefund = trxData[i][5]; 
       break;
    }
  }

  keuSheet.appendRow([
      'REFUND-' + Date.now(), new Date(), 
      'Pengeluaran', 'Retur Penjualan', 
      nominalRefund, `Retur: ${produkNama} (${idTrx})`
  ]);
  catatLog("HAPUS_TRANSAKSI", "Menghapus Transaksi ID: " + idTrx);

  SpreadsheetApp.flush(); // Tambahkan ini

  return "Berhasil Retur/Hapus";
}

// ==========================================
// 7. PEMBELIAN & SUPPLIER
// ==========================================

function simpanPembelianBulk(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetBeli = ss.getSheetByName('PEMBELIAN');
  const sheetProd = ss.getSheetByName('PRODUK');
  const sheetKeu = ss.getSheetByName('KEUANGAN');
   
  const idBeliMaster = 'BELI-' + Date.now();
  const waktu = new Date();
  const prodData = sheetProd.getDataRange().getValues();
   
  let summaryItem = [];

  data.items.forEach(item => {
    sheetBeli.appendRow([
      idBeliMaster, 
      waktu, 
      data.supplier, 
      item.produk, 
      item.qty, 
      item.total, 
      'Tunai'
    ]);

    for (let i = 1; i < prodData.length; i++) {
      if (prodData[i][1] == item.produk) {
        
        // --- LOGIKA AVERAGE PRICE ---
        let curIsi = Number(prodData[i][4]) || 0;       
        let curHargaBeli = Number(prodData[i][3]) || 0; 
        
        // [FIX] Definisikan curKosong sebelum dipakai
        let curKosong = Number(prodData[i][5]) || 0;    

        let qtyBeli = Number(item.qty);
        let hargaBeliBaru = Number(item.hargaSatuan);   

        let nilaiAsetLama = curIsi * curHargaBeli;
        let nilaiAsetBaru = qtyBeli * hargaBeliBaru;

        let totalStok = curIsi + qtyBeli;
        let averagePrice = curHargaBeli; 

        if (totalStok > 0) {
            averagePrice = (nilaiAsetLama + nilaiAsetBaru) / totalStok;
            averagePrice = Math.round(averagePrice); 
        }

        sheetProd.getRange(i + 1, 4).setValue(averagePrice); 
        sheetProd.getRange(i + 1, 5).setValue(totalStok);   
        
        if(item.isTukar) {
           sheetProd.getRange(i + 1, 6).setValue(curKosong - Number(item.qty));
        }
        break;
      }
    }
    summaryItem.push(`${item.produk} (x${item.qty})`);
  });

  sheetKeu.appendRow([
    'OUT-' + Date.now(), 
    waktu, 
    'Pengeluaran', 
    'Pembelian Stok', 
    data.grandTotal, 
    `Beli Stok: ${summaryItem.join(', ')}`
  ]);

  return "Stok Berhasil Ditambahkan & Harga Rata-rata Diupdate!";
}

function tambahSupplier(form) {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SUPPLIER').appendRow(['SUP-' + Date.now(), form.nama, form.hp, form.alamat]);
}

// Fungsi Simpan Beli Satuan (Legacy - Tetap Disimpan)
function simpanPembelian(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const prodSheet = ss.getSheetByName('PRODUK');
  
  ss.getSheetByName('PEMBELIAN').appendRow(['BELI-' + Date.now(), new Date(), data.supplier, data.produk, data.qty, data.total, data.metode]);
  
  const prodData = prodSheet.getDataRange().getValues();
  for (let i = 1; i < prodData.length; i++) {
    if (prodData[i][1] == data.produk) {
      let curIsi = Number(prodData[i][4]);
      let curKosong = Number(prodData[i][5]);
      
      prodSheet.getRange(i + 1, 5).setValue(curIsi + Number(data.qty)); 
      if(data.isTukar) {
        prodSheet.getRange(i + 1, 6).setValue(curKosong - Number(data.qty)); 
      }
      break;
    }
  }
  
  ss.getSheetByName('KEUANGAN').appendRow(['OUT-' + Date.now(), new Date(), 'Pengeluaran', 'Pembelian Stok', data.total, `Beli ${data.produk}`]);
}

function getRiwayatPembelian() {
  const data = getData('PEMBELIAN');
  let grouped = {};

  data.forEach(row => {
    let id = row[0];
    let waktuStr = row[1] instanceof Date ? row[1].toISOString() : row[1];

    if (!grouped[id]) {
      grouped[id] = {
        id: id,
        waktu: waktuStr,
        pelanggan: row[2], 
        totalBayar: 0,
        items: []
      };
    }

    grouped[id].items.push({
      produk: row[3],
      qty: row[4],
      hargaTotal: row[5],
      tipe: 'Stok Masuk', 
      status: 'Sukses' 
    });
     
    grouped[id].totalBayar += Number(row[5]);
  });

  return Object.values(grouped).sort((a, b) => new Date(b.waktu) - new Date(a.waktu)).slice(0, 50);
}

// ==========================================
// 8. PELANGGAN
// ==========================================

function simpanPelanggan(form) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PELANGGAN');
   
  if(form.id) { 
    const data = sheet.getDataRange().getValues();
    for(let i=1; i<data.length; i++) {
      if(data[i][0] == form.id) {
        sheet.getRange(i+1, 2, 1, 4).setValues([[form.nama, form.pt, form.hp, form.alamat]]);
        return "Data Pelanggan Diupdate";
      }
    }
  }
   
  sheet.appendRow(['CUST-' + Date.now(), form.nama, form.pt, form.hp, form.alamat]);
  return "Pelanggan Baru Disimpan";
}

function hapusPelanggan(id) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PELANGGAN');
  const data = sheet.getDataRange().getValues();
  for(let i=1; i<data.length; i++) {
    if(data[i][0] == id) { 
      sheet.deleteRow(i+1); 
      return "Pelanggan Dihapus";
    }
  }
}

function getListPelanggan() {
  return getData('PELANGGAN'); 
}

// ==========================================
// 9. KEUANGAN & KATEGORI
// ==========================================

function getKategori() {
  return getData('KATEGORI').map(r => r[0]);
}

function tambahKategori(nama) {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('KATEGORI').appendRow([nama]);
}

function simpanKeuangan(form) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('KEUANGAN');
  const tglInput = new Date(form.tanggal);

  if(sheet.getLastColumn() < 7) {
     sheet.getRange(1, 7).setValue('Akun');
  }

  // LOGIKA BARU
  const newId = 'MANUAL-' + Date.now();
  sheet.appendRow([
      newId, 
      tglInput, 
      form.jenis, 
      form.kategori, 
      form.nominal, 
      form.keterangan,
      form.akun 
  ]);
   
  return { status: 'success', data: { id: newId, ...form } };
}

function simpanAkunBaru(form) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AKUN_KAS');
  
  // --- CEK: APAKAH INI EDIT DATA? (JIKA ADA ID) ---
  if (form.id && form.id !== "") {
     const data = sheet.getDataRange().getValues();
     
     for(let i=1; i<data.length; i++) {
        if(data[i][0] == form.id) {
           // Simpan data lama untuk laporan log yang detail
           let namaLama = data[i][1];
           let saldoLama = data[i][4];

           // Lakukan Update Cell
           sheet.getRange(i+1, 2).setValue(form.nama);
           sheet.getRange(i+1, 3).setValue("'" + form.norek);
           sheet.getRange(i+1, 4).setValue(form.tipe);
           sheet.getRange(i+1, 5).setValue(form.saldo);
           
           // [PENTING] CATAT LOG EDIT
           catatLog("EDIT_AKUN_KAS", `Update Akun: ${namaLama} | Saldo: ${saldoLama} -> ${form.saldo}`);
           
           return "Data Akun Berhasil Diupdate!";
        }
     }
  }

  // --- JIKA TIDAK ADA ID, BERARTI TAMBAH BARU ---
  const newId = 'ACC-' + Date.now();
  sheet.appendRow([newId, form.nama, "'" + form.norek, form.tipe, form.saldo]); 
  
  // [PENTING] CATAT LOG BARU
  catatLog("TAMBAH_AKUN_KAS", "Menambah Akun Baru: " + form.nama + " | Saldo Awal: " + form.saldo);

  return "Akun Berhasil Ditambahkan!";
}

function hapusAkun(id) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AKUN_KAS');
  const data = sheet.getDataRange().getValues();
  
  for(let i=1; i<data.length; i++) {
     if(data[i][0] == id) {
        // 1. Ambil nama akun dulu sebelum dihapus agar log-nya jelas
        let namaAkun = data[i][1]; 
        let saldoTerakhir = data[i][4];

        // 2. Hapus Baris
        sheet.deleteRow(i+1);
        
        // 3. [PENTING] CATAT KE LOG AKTIVITAS
        catatLog("HAPUS_AKUN_KAS", "Menghapus Akun: " + namaAkun + " | ID: " + id + " | Saldo Terakhir: " + saldoTerakhir);

        return "Akun Dihapus.";
     }
  }
  return "ID Akun tidak ditemukan.";
}

function prosesTransferSaldo(form) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('KEUANGAN');
  const waktu = new Date();
  const idTrx = 'TRF-' + Date.now();

  sheet.appendRow([
     idTrx + '-OUT', waktu, 'Pengeluaran', 'Transfer Keluar', form.nominal, 
     `Transfer ke ${form.akunTujuan} (${form.ket})`, form.akunAsal
  ]);

  sheet.appendRow([
     idTrx + '-IN', waktu, 'Pemasukan', 'Transfer Masuk', form.nominal, 
     `Terima dari ${form.akunAsal} (${form.ket})`, form.akunTujuan
  ]);

  return "Transfer Berhasil!";
}

function hapusKeuangan(id) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('KEUANGAN');
  const data = sheet.getDataRange().getValues();
   
  if(!String(id).includes('MANUAL')) {
     throw new Error("Data sistem (Otomatis) tidak boleh dihapus dari sini!");
  }

  for(let i = 1; i < data.length; i++) {
    if(data[i][0] == id) {
       sheet.deleteRow(i+1);
       return "Data Dihapus";
    }
  }
  throw new Error("ID tidak ditemukan");
}

// [CODE.GS] Fungsi Get Laporan Detail (Revisi Pemisahan Retur)
function getLaporanDetail(startStr, endStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
   
  const startDate = new Date(startStr); startDate.setHours(0,0,0,0);
  const endDate = new Date(endStr); endDate.setHours(23,59,59,999);

  const prodData = getData('PRODUK');
  const trxData = getData('TRANSAKSI');
  const keuData = getData('KEUANGAN');

  let laporan = {
    pendapatan: 0,
    hpp: 0,
    pengeluaran: 0, // Khusus Biaya Operasional (Gaji, Listrik, dll)
    totalRetur: 0,  // [BARU] Khusus Retur Penjualan
    pembelian: 0,
    nilaiStok: 0,      
    asetFisik: 0,      
    asetDagang: 0,     
    labaKotor: 0,
    labaBersih: 0,
    metodeBayar: {},
    listPengeluaran: []
  };

  // 1. MAPPING DATA PRODUK
  let mapProduk = {};
  prodData.forEach(p => {
    let hppIsi = Number(p[3]) || 0;
    let hppTabung = Number(p[9]) || 0;
    let stokIsi = Number(p[4]) || 0;
    let stokKosong = Number(p[5]) || 0;
    
    mapProduk[p[1]] = { hppIsi: hppIsi, hppTabung: hppTabung, stokIsi: stokIsi };
    
    let nilaiCangkang = (stokIsi + stokKosong) * hppTabung;
    let nilaiIsi = stokIsi * hppIsi;
    
    laporan.asetFisik += nilaiCangkang;
    laporan.asetDagang += nilaiIsi;
    laporan.nilaiStok += (nilaiCangkang + nilaiIsi);
  });

  // 2. HITUNG TRANSAKSI (PEMASUKAN & HPP AWAL)
  trxData.forEach(row => {
    let tgl = new Date(row[1]);
    if (tgl >= startDate && tgl <= endDate) {
        let produk = row[3]; 
        let qty = Number(row[4]);
        let total = Number(row[5]);
        let tipe = String(row[6]).toLowerCase();
        
        let hppSatuan = 0;
        if(mapProduk[produk]) {
            if(tipe.includes('baru')) {
                hppSatuan = mapProduk[produk].hppIsi + mapProduk[produk].hppTabung;
            } else {
                hppSatuan = mapProduk[produk].hppIsi;
            }
        }
        
        laporan.pendapatan += total;
        laporan.hpp += (hppSatuan * qty);

        // Hitung Metode Bayar
        let metode = row[8];
        if(!laporan.metodeBayar[metode]) laporan.metodeBayar[metode] = 0;
        laporan.metodeBayar[metode] += total;
    }
  });

  // 3. HITUNG PENGELUARAN & RETUR (LOGIKA UTAMA)
  keuData.forEach(row => {
    let tgl = new Date(row[1]);
    if (tgl >= startDate && tgl <= endDate) {
      
      // -- SKENARIO 1: RETUR PENJUALAN --
      if (row[3] === 'Retur Penjualan') {
         let nilaiRefund = Number(row[4]) || 0;
         let nilaiHPPRetur = Number(row[6]) || 0;
         
         // A. Pendapatan JANGAN dikurangi (agar Omzet tetap utuh sesuai struk awal)
         // laporan.pendapatan -= nilaiRefund; // (Dihapus/Disable)

         // B. HPP DIKURANGI (Karena barang balik ke stok, HPP batal keluar)
         laporan.hpp -= nilaiHPPRetur; 

         // C. Masukkan ke wadah khusus RETUR (bukan pengeluaran operasional)
         laporan.totalRetur += nilaiRefund; 

         // D. Tetap masukkan ke tabel rincian (opsional, agar bisa dicek)
         laporan.listPengeluaran.push({
            id: row[0],
            tanggal: row[1],
            kategori: 'Retur Penjualan',
            ket: row[5],
            jumlah: nilaiRefund
         });

      // -- SKENARIO 2: PENGELUARAN LAIN / PEMBELIAN STOK --
      } else if (row[2] === 'Pengeluaran') {

        if (row[3] === 'Pembelian Stok') {
           // Pembelian stok tidak mengurangi laba rugi langsung (masuk aset)
           laporan.pembelian += Number(row[4]);
        } else {
           // Ini Biaya Operasional (Listrik, Gaji, Makan, dll)
           laporan.pengeluaran += Number(row[4]);
           
           laporan.listPengeluaran.push({
              id: row[0],
              tanggal: row[1],
              kategori: row[3],
              ket: row[5],
              jumlah: Number(row[4])
           });
        }
      }
    }
  });

  // 4. HITUNG LABA BERSIH (Final Calculation)
  // Laba Kotor = Omzet - HPP (yg sudah dikurangi retur barang)
  laporan.labaKotor = laporan.pendapatan - laporan.hpp;
  
  // Laba Bersih = Laba Kotor - Biaya Ops - Uang Kembali (Retur)
  laporan.labaBersih = laporan.labaKotor - laporan.pengeluaran - laporan.totalRetur;

  laporan.daftarAkun = getDaftarAkun();

  return laporan;
}
// ==========================================
// 10. KARYAWAN & PAYROLL
// ==========================================

function simpanKaryawan(form) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('KARYAWAN');
  const data = sheet.getDataRange().getValues();
   
  const FOLDER_ID = '15hiLtvusofF2OJpXVq8lJkePbmqVIuPM'; 

  let fotoUrl = form.fotoLama || ''; 
   
  if (form.foto && form.foto.data) {
    try {
      const decoded = Utilities.base64Decode(form.foto.data);
      const blob = Utilities.newBlob(decoded, form.foto.mimeType, 'KRY-' + Date.now());
      const folder = DriveApp.getFolderById(FOLDER_ID);
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      fotoUrl = "https://drive.google.com/uc?export=view&id=" + file.getId();
    } catch (e) {
      console.log("Upload Error: " + e.message);
    }
  }

    const rowData = [
    form.nama, form.tmpLahir, form.tglLahir, form.gender, 
    "'" + form.noId, form.tipeId, form.email, form.alamatKtp, form.alamatDom,
    form.daruratNama, "'" + form.daruratTelp, form.gaji, form.bonus, fotoUrl, 'Aktif', 
    form.tglMasuk 
  ];
   
  if(form.id) { 
    for(let i=1; i<data.length; i++) {
      if(data[i][0] == form.id) {
        const range = sheet.getRange(i+1, 2, 1, 16); 
        range.setValues([rowData]);
        return "Data Karyawan Berhasil Diperbarui";
      }
    }
  }
   
  const newId = 'KRY-' + Date.now();
  sheet.appendRow([newId, ...rowData]); 
  return "Karyawan Baru Berhasil Disimpan";
}

function hapusKaryawan(id) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('KARYAWAN');
  const data = sheet.getDataRange().getValues();
  for(let i=1; i<data.length; i++) {
    if(data[i][0] == id) { sheet.deleteRow(i+1); return; }
  }
}

function bayarCicilanManual(form) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetKasbon = ss.getSheetByName('KASBON');
  const sheetRiwayat = ss.getSheetByName('RIWAYAT_BAYAR_KASBON');
  const sheetKeu = ss.getSheetByName('KEUANGAN');
   
  if (!form || !form.idKasbon || !form.nominal) {
    throw new Error("Data pembayaran tidak lengkap.");
  }

  const headers = sheetKasbon.getRange(1, 1, 1, sheetKasbon.getLastColumn()).getValues()[0];
  const cariIndex = (nama) => headers.findIndex(h => String(h).toLowerCase() === nama.toLowerCase());

  const idxID = cariIndex("ID_Kasbon");
  const idxTotal = cariIndex("Nominal");        
  const idxSudahBayar = cariIndex("Sudah_Bayar"); 
  const idxStatus = cariIndex("Status_Lunas");

  if (idxID === -1 || idxTotal === -1 || idxSudahBayar === -1) {
    throw new Error("Struktur Kolom Sheet KASBON tidak sesuai!");
  }

  const dataKasbon = sheetKasbon.getDataRange().getValues();
  let foundRowIndex = -1;
  let totalPinjam = 0;
  let sudahBayarLama = 0;
   
  for (let i = 1; i < dataKasbon.length; i++) {
    if (String(dataKasbon[i][idxID]) === String(form.idKasbon)) {
      foundRowIndex = i + 1; 
      totalPinjam = Number(dataKasbon[i][idxTotal]) || 0;
      sudahBayarLama = Number(dataKasbon[i][idxSudahBayar]) || 0;
      break;
    }
  }

  if (foundRowIndex === -1) throw new Error("ID Kasbon tidak ditemukan.");

  const bayarIni = Number(form.nominal);
  const sudahBayarBaru = sudahBayarLama + bayarIni;
  const sisaHutangBaru = totalPinjam - sudahBayarBaru; 
  const statusLunas = sisaHutangBaru <= 0 ? 'Lunas' : 'Belum Lunas';

  sheetKasbon.getRange(foundRowIndex, idxSudahBayar + 1).setValue(sudahBayarBaru);
  sheetKasbon.getRange(foundRowIndex, idxStatus + 1).setValue(statusLunas);

  const timeNow = new Date();
  const akunTujuan = form.akun || 'Kas Tunai'; 
  const infoMetode = `Via ${akunTujuan}`; 

  sheetRiwayat.appendRow([
      'PAY-' + Date.now(),             
      form.idKasbon,                   
      timeNow,                          
      bayarIni,                        
      'Manual',                        
      infoMetode 
  ]);

  sheetKeu.appendRow([
     'INC-' + Date.now(),    
     timeNow,                
     'Pemasukan',            
     'Cicilan Kasbon', 
     bayarIni, 
     `Cicilan Hutang ${form.namaKaryawan || ''} (${infoMetode})`, 
     akunTujuan              
  ]);
   
  SpreadsheetApp.flush(); 
  return `Pembayaran masuk ke ${akunTujuan}. Sisa Hutang: ${sisaHutangBaru.toLocaleString('id-ID')}`;
}

function getDetailHistoryKasbon(idKasbon) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('RIWAYAT_BAYAR_KASBON');
  if (!sheet) return [];
   
  const data = sheet.getDataRange().getValues();
  const targetID = String(idKasbon).trim();
  let history = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    let rowID = String(row[1]).trim();

    if (rowID === targetID) {
       let tglRaw = row[2];
       let tglDisplay = tglRaw;
       
       if (tglRaw instanceof Date) {
          tglDisplay = Utilities.formatDate(tglRaw, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd HH:mm");
       } else {
          tglDisplay = String(tglRaw);
       }

       history.push({
          tanggal: tglDisplay,
          nominal: Number(row[3]), 
          tipe: row[4],
          ket: row[5]
       });
    }
  }

  return history.sort((a,b) => {
      if (a.tanggal < b.tanggal) return 1;
      if (a.tanggal > b.tanggal) return -1;
      return 0;
  });
}

function simpanKasbon(form) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetKsb = ss.getSheetByName('KASBON');
  const sheetKeu = ss.getSheetByName('KEUANGAN');

  const tenor = Number(form.tenor) || 1;
  const nominal = Number(form.nominal);
  const angsuran = Math.ceil(nominal / tenor);
  const idKasbon = 'KSB-' + Date.now();
  const waktu = new Date();

  const ketLengkap = `${form.ket} (Via ${form.akun})`;

  sheetKsb.appendRow([
      idKasbon, 
      waktu, 
      form.nama, 
      nominal, 
      ketLengkap,      
      'Belum Lunas', 
      0, 
      tenor, 
      angsuran 
  ]);

  sheetKeu.appendRow([
      'OUT-' + idKasbon,        
      waktu, 
      'Pengeluaran',            
      'Kasbon Karyawan',        
      nominal, 
      `Pencairan Kasbon: ${form.nama} (${form.ket})`, 
      form.akun                 
  ]);

  return "Kasbon berhasil disimpan & Saldo terpotong.";
}

function getDataPayroll() {
  const karyawan = getData('KARYAWAN');
  const kasbonData = getData('KASBON');
  if (!karyawan || karyawan.length === 0) return [];

  let result = karyawan.map(k => {
    let nama = k[1]; 
    let gaji = Number(k[12]) || 0; 
    let bonusSet = Number(k[13]) || 0; 
    
    let totalHutang = 0;
    let tagihanBulanIni = 0;
    let infoTenorList = [];

    if (kasbonData && kasbonData.length > 0) {
        kasbonData.forEach(ksb => {
          if(ksb[2] === nama && String(ksb[5]).includes('Belum Lunas')) {
            let nominalUtang = Number(ksb[3]);
            let sudahBayar = Number(ksb[6]) || 0;
            let sisa = nominalUtang - sudahBayar;
            
            totalHutang += sisa;

            let tenorTotal = Number(ksb[7]) || 1; 
            let angsuran = Number(ksb[8]); 
            if(!angsuran || angsuran === 0) angsuran = sisa;

            let cicilanKe = Math.floor(sudahBayar / angsuran) + 1;
            if(cicilanKe > tenorTotal) cicilanKe = tenorTotal;

            infoTenorList.push(`(${cicilanKe}/${tenorTotal})`);
            
            let harusBayar = Math.min(angsuran, sisa);
            tagihanBulanIni += harusBayar;
          }
        });
    }

    return {
      id: k[0],
      nama: nama,
      gaji: gaji,
      bonus: bonusSet,
      sisaHutang: totalHutang, 
      kasbonPotongan: tagihanBulanIni,
      infoTenor: infoTenorList.join(', '),
      total: gaji + bonusSet - tagihanBulanIni
    };
  });
   
  return result;
}

function getDataKasbonFull() {
  return getData('KASBON');
}

function getRiwayatGajiByPeriode(periode) { 
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetRiwayat = ss.getSheetByName('RIWAYAT_GAJI');
  const sheetKaryawan = ss.getSheetByName('KARYAWAN');
  const rawKaryawan = sheetKaryawan.getDataRange().getValues().slice(1); 
  const kasbonData = getData('KASBON'); 

  let year = parseInt(periode.split('-')[0]);
  let month = parseInt(periode.split('-')[1]) - 1; 
   
  let startMonth = new Date(year, month, 1);
  let endMonth = new Date(year, month + 1, 0); 
   
  let totalHariKerjaSebulan = hitungHariKerja(startMonth, endMonth);

  const riwayatRaw = sheetRiwayat ? sheetRiwayat.getDataRange().getDisplayValues() : [];
  let dataSavedMap = {};
  let targetPeriode = String(periode).trim();

  for(let i=1; i<riwayatRaw.length; i++) {
     let rowP = String(riwayatRaw[i][1]).trim();
     if(rowP === targetPeriode) {
        dataSavedMap[riwayatRaw[i][3]] = {
           id: riwayatRaw[i][0],
           periode: rowP,
           nama: riwayatRaw[i][3],
           gaji: parseDuit(riwayatRaw[i][4]), 
           bonus: parseDuit(riwayatRaw[i][5]),
           kasbon: parseDuit(riwayatRaw[i][6]),
           total: parseDuit(riwayatRaw[i][7]),
           status: riwayatRaw[i][8],
           kasbonBaru: parseDuit(riwayatRaw[i][9]) 
        };
     }
  }

  let hasilFinal = rawKaryawan.map(k => {
      let nama = k[1];
      
      if (dataSavedMap[nama]) {
          return dataSavedMap[nama];
      } 
      
      let gajiPokokDB = Number(k[12]) || 0;
      let bonusDB = Number(k[13]) || 0;
      let tglMasukDB = k[16]; 
      
      let gajiFinal = gajiPokokDB;
      
      if (tglMasukDB instanceof Date) {
          if (tglMasukDB > startMonth && tglMasukDB <= endMonth) {
              let hariKerjaDia = hitungHariKerja(tglMasukDB, endMonth);
              let rasio = hariKerjaDia / totalHariKerjaSebulan;
              gajiFinal = Math.floor(gajiPokokDB * rasio); 
              gajiFinal = Math.floor(gajiFinal / 100) * 100;
          }
      }
      
      let tagihanKasbon = 0;
      kasbonData.forEach(ksb => {
          if(ksb[2] === nama && String(ksb[5]).includes('Belum')) {
             let sisa = Number(ksb[3]) - (Number(ksb[6]) || 0);
             let angsuran = Number(ksb[8]) || sisa;
             tagihanKasbon += Math.min(angsuran, sisa);
          }
      });

      return {
         id: 'DRAFT-' + Date.now(),
         periode: targetPeriode,
         nama: nama,
         gaji: gajiFinal, 
         bonus: bonusDB,
         kasbon: tagihanKasbon,
         total: gajiFinal + bonusDB - tagihanKasbon,
         status: 'Belum Digaji'
      };
  });

  return hasilFinal;
}

// Helper untuk membersihkan "Rp 3.000.000" jadi angka 3000000
function parseDuit(val) {
   if(!val) return 0;
   let bersih = String(val).replace(/[^0-9,-]+/g,""); 
   return Number(bersih) || 0;
}

function simpanGajiBulanan(periode, listGaji, akunBayar) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetRiwayat = ss.getSheetByName('RIWAYAT_GAJI');
  const sheetKasbonHistory = ss.getSheetByName('RIWAYAT_BAYAR_KASBON'); 
  const keuSheet = ss.getSheetByName('KEUANGAN');
  const kasbonSheet = ss.getSheetByName('KASBON');
  const kasbonData = kasbonSheet.getDataRange().getValues();
  const waktu = new Date();
   
  const periodeClean = String(periode).trim();

  // 1. HAPUS DATA LAMA 
  const dataLama = sheetRiwayat.getDataRange().getDisplayValues();
  const namaYangDiproses = listGaji.map(x => x.nama);

  for (let i = dataLama.length - 1; i >= 1; i--) {
      let rowPeriode = String(dataLama[i][1]).trim();
      let rowNama = String(dataLama[i][3]).trim();
      if (rowPeriode === periodeClean && namaYangDiproses.includes(rowNama)) {
          sheetRiwayat.deleteRow(i + 1);
      }
  }

  // 2. SIMPAN DATA BARU
  let totalKeluarSesiIni = 0;

  listGaji.forEach(k => {
      // A. Kasbon Baru
      let nominalKasbonBaru = Number(k.kasbonBaru) || 0;
      if (nominalKasbonBaru > 0) {
          kasbonSheet.appendRow(['KSB-' + Date.now(), waktu, k.nama, nominalKasbonBaru, k.ketKasbonBaru || 'Kasbon via Payroll', 'Belum Lunas', 0, 1, nominalKasbonBaru]);
      }

      // B. Potongan Kasbon
      let bayarBulanIni = Number(k.kasbonPotongan) || 0; 

      if(bayarBulanIni > 0) {
          for(let i=1; i<kasbonData.length; i++) {
            if(kasbonData[i][2] == k.nama && String(kasbonData[i][5]).includes('Belum') && bayarBulanIni > 0) {
               let sisaHutangDb = Number(kasbonData[i][3]) - (Number(kasbonData[i][6]) || 0);
               let alokasi = Math.min(sisaHutangDb, bayarBulanIni);
               let newSudah = (Number(kasbonData[i][6]) || 0) + alokasi;
               
               kasbonSheet.getRange(i+1, 7).setValue(newSudah); 
               
               if(newSudah >= Number(kasbonData[i][3])) {
                   kasbonSheet.getRange(i+1, 6).setValue('Lunas');
               }

               sheetKasbonHistory.appendRow(['AUTO-' + Date.now(), kasbonData[i][0], waktu, alokasi, 'Potong Gaji', 'Payroll ' + periode]);
               
               bayarBulanIni -= alokasi;
            }
          }
      }

      // C. Simpan ke RIWAYAT GAJI
      let potonganLain = Number(k.potonganLain) || 0;
      let totalPotonganDicatat = (Number(k.kasbonPotongan) || 0) + potonganLain;

      sheetRiwayat.appendRow([
          'PAY-' + Date.now(), 
          periodeClean, 
          waktu, 
          k.nama, 
          k.gaji, 
          Number(k.bonus), 
          totalPotonganDicatat, 
          k.total, 
          'Sukses',
          nominalKasbonBaru 
      ]);
      
      totalKeluarSesiIni += Number(k.total);
  });

  if(totalKeluarSesiIni > 0) {
      const akunFinal = akunBayar || 'Kas Tunai (Laci)';
      
      keuSheet.appendRow([
          'PAYROLL-' + Date.now(), 
          waktu, 
          'Pengeluaran', 
          'Gaji Karyawan', 
          totalKeluarSesiIni, 
          `Payroll ${periodeClean} (${listGaji.length} Org)`, 
          akunFinal 
      ]);
  }
   
  SpreadsheetApp.flush(); 
  return "Gaji berhasil dicairkan! Kasbon & Saldo terupdate.";
}

function hitungHariKerja(tglMulai, tglSelesai) {
  let count = 0;
  let curDate = new Date(tglMulai.getTime());
   
  while (curDate <= tglSelesai) {
    let dayOfWeek = curDate.getDay();
    if(dayOfWeek !== 0) { 
       count++;
    }
    curDate.setDate(curDate.getDate() + 1);
  }
  return count;
}

// ==========================================
// 11. DASHBOARD & REPORTING
// ==========================================

// [CODE.GS] UPDATE FUNGSI getDashboardRealtime
function getDashboardRealtime(startStr, endStr) {
  let startDate, endDate;
  
  if(startStr && endStr) {
    startDate = new Date(startStr);
    endDate = new Date(endStr);
    endDate.setHours(23, 59, 59);
  } else {
    let now = new Date();
    startDate = new Date(now.getFullYear(), now.getMonth(), 1); 
    endDate = new Date(now.getFullYear(), now.getMonth() + 1, 0, 23, 59, 59); 
  }

  let stats = {
    pendapatan: 0,
    pengeluaran: 0,
    labaBersih: 0,
    pesanan: 0,
    totalPelanggan: 0,
    produkTerjual: {},
    chartSales: new Array(31).fill(0),
    chartExpense: new Array(31).fill(0),
    qtyRefill: 0,
    rataRataTransaksi: 0,
    retur: 0 
  };
  
  // Helper mapping tanggal ke index array (0-30)
  let mapDateIndex = {};
  for(let i=0; i<=30; i++) {
     let d = new Date(startDate);
     d.setDate(d.getDate() + i);
     let key = Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
     mapDateIndex[key] = i;
  }

  const rawTrx = getData('TRANSAKSI');
  const rawKeu = getData('KEUANGAN');
  const rawPel = getData('PELANGGAN');
  const rawAkun = getDaftarAkun();

  stats.totalPelanggan = rawPel.length;

  // A. LOOP TRANSAKSI
  if (rawTrx && rawTrx.length > 0) {
      rawTrx.forEach(row => {
          let tgl = new Date(row[1]); 
          if(tgl >= startDate && tgl <= endDate) {
              let produk = row[3];
              let qty = Number(row[4]) || 0;
              let total = Number(row[5]) || 0;
              let tipe = String(row[6]).toLowerCase();
              let status = String(row[10]); // Ambil Status

              // [FIX PENTING] 
              // Hapus pengecekan if(!status.includes('Retur')) 
              // Agar transaksi yang statusnya 'Retur' TETAP DIHITUNG pemasukannya.
              
              let diskon = (row[11]) ? Number(row[11]) : 0; 
              if (isNaN(diskon)) diskon = 0;

              // Kita hitung semua transaksi (kecuali jika ada status 'Batal' mungkin perlu di skip, tapi 'Retur' tetap hitung)
              if(!status.toLowerCase().includes('batal')) {
                  stats.pesanan++;
                  
                  let pendapatanBersih = total - diskon;
                  stats.pendapatan += pendapatanBersih; // 52.500 tetap masuk sini
                  
                  if(!stats.produkTerjual[produk]) stats.produkTerjual[produk] = 0;
                  stats.produkTerjual[produk] += qty;

                  if(tipe && tipe.includes('refill')) {
                    stats.qtyRefill += qty;
                  }
                  
                  let key = Utilities.formatDate(tgl, Session.getScriptTimeZone(), "yyyy-MM-dd");
                  if (mapDateIndex.hasOwnProperty(key)) {
                    let i = mapDateIndex[key];
                    stats.chartSales[i] += pendapatanBersih;
                  }
              }
          }
      });
  }

  // B. LOOP KEUANGAN
  if (rawKeu && rawKeu.length > 0) {
      rawKeu.forEach(row => {
          let tgl = new Date(row[1]);
          if(tgl >= startDate && tgl <= endDate) {
            let jenis = row[2];
            let kategori = String(row[3]).toLowerCase();
            let nominal = Number(row[4]) || 0; 
            let key = Utilities.formatDate(tgl, Session.getScriptTimeZone(), "yyyy-MM-dd");

            if(jenis === 'Pengeluaran') {
                stats.pengeluaran += nominal; // 17.500 masuk sini sebagai pengurang
                
                if(kategori.includes('retur')) {
                  stats.retur += nominal;
                }

                if (mapDateIndex.hasOwnProperty(key)) {
                  let i = mapDateIndex[key];
                  stats.chartExpense[i] += nominal;
                }
            }
          }
      });
  }

  if(stats.pesanan > 0) {
      stats.rataRataTransaksi = stats.pendapatan / stats.pesanan;
  }

  return {
      stats: stats,
      accounts: rawAkun,
      periode: `${Utilities.formatDate(startDate, Session.getScriptTimeZone(), "dd MMM")} - ${Utilities.formatDate(endDate, Session.getScriptTimeZone(), "dd MMM yyyy")}`
  };
}

function getNotifikasiData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const today = new Date();
  let notifikasiList = [];
   
  // 1. CEK HUTANG PELANGGAN
  const trxSheet = ss.getSheetByName('TRANSAKSI');
  if (trxSheet) {
    const dataTrx = trxSheet.getDataRange().getValues();
    for (let i = 1; i < dataTrx.length; i++) {
      let row = dataTrx[i];
      if (row[8] === 'Hutang' && row[10] === 'Belum Lunas' && row[9]) {
        let tglJatuhTempo = new Date(row[9]);
        let diffDays = Math.ceil((tglJatuhTempo - today) / (1000 * 60 * 60 * 24)); 
        
        if (diffDays <= 3) { 
          let statusWaktu = (diffDays < 0) ? 'Telat ' + Math.abs(diffDays) + ' hari' : 'H-' + diffDays;
          notifikasiList.push({
            tipe: 'hutang',
            judul: 'Tagihan Pelanggan',
            pesan: `${row[2]} jatuh tempo (${statusWaktu})`,
            waktu: Utilities.formatDate(tglJatuhTempo, Session.getScriptTimeZone(), 'dd/MM/yyyy')
          });
        }
      }
    }
  }

  // 2. CEK KASBON KARYAWAN
  const kasbonSheet = ss.getSheetByName('KASBON');
  if (kasbonSheet) {
    const dataKasbon = kasbonSheet.getDataRange().getValues();
    for (let i = 1; i < dataKasbon.length; i++) {
      let row = dataKasbon[i];
      if (row[5] !== 'Lunas' && row[4]) {
        let tglJatuhTempo = new Date(row[4]);
        let diffDays = Math.ceil((tglJatuhTempo - today) / (1000 * 60 * 60 * 24));
        
        if (diffDays <= 3) {
          let statusWaktu = (diffDays < 0) ? 'Telat ' + Math.abs(diffDays) + ' hari' : 'H-' + diffDays;
          notifikasiList.push({
            tipe: 'kasbon',
            judul: 'Kasbon Karyawan',
            pesan: `${row[1]} belum lunas (${statusWaktu})`,
            waktu: Utilities.formatDate(tglJatuhTempo, Session.getScriptTimeZone(), 'dd/MM/yyyy')
          });
        }
      }
    }
  }

  // 3. CEK STOK MENIPIS
  const prodSheet = ss.getSheetByName('PRODUK');
  if (prodSheet) {
    const dataProd = prodSheet.getDataRange().getValues();
    const BATAS_MINIM = 5; 

    for (let i = 1; i < dataProd.length; i++) {
      let row = dataProd[i];
      let namaProduk = row[1];
      let stokReady = Number(row[4]);

      if (stokReady <= BATAS_MINIM) {
        let pesanStok = (stokReady <= 0) ? 'HABIS!' : 'Sisa ' + stokReady;
        notifikasiList.push({
          tipe: 'stok',
          judul: 'Stok Kritis',
          pesan: `${namaProduk} segera habis (${pesanStok})`,
          waktu: 'Cek Gudang'
        });
      }
    }
  }
   
  return notifikasiList;
}

// ==========================================
// 12. LOG & EXPORT PDF
// ==========================================

function catatLog(aksi, detail) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetLog = ss.getSheetByName('LOG_AKTIVITAS');
     
    if (!sheetLog) {
      sheetLog = ss.insertSheet('LOG_AKTIVITAS');
      sheetLog.appendRow(['TIMESTAMP', 'USER / EMAIL', 'JENIS AKSI', 'DETAIL AKTIVITAS']);
      sheetLog.setFrozenRows(1);
      sheetLog.getRange("A1:D1").setFontWeight("bold").setBackground("#ddd");
      sheetLog.setColumnWidth(1, 150); 
      sheetLog.setColumnWidth(4, 400); 
    }
     
    var waktu = new Date(); 
    var user = Session.getActiveUser().getEmail(); 
    if (!user) user = 'Guest/Admin'; 

    sheetLog.appendRow([waktu, user, aksi, detail]);
     
  } catch (e) {
    Logger.log("Gagal mencatat log: " + e.message);
  }
}

function getLogAktivitas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('LOG_AKTIVITAS');
  if (!sheet) return [];
   
  var lastRow = sheet.getLastRow();
  var startRow = Math.max(2, lastRow - 50); 
  if (lastRow < 2) return []; 
   
  var data = sheet.getRange(startRow, 1, lastRow - startRow + 1, 4).getValues();
   
  return data.reverse().map(function(row) {
    return {
      waktu: Utilities.formatDate(new Date(row[0]), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss"),
      user: row[1],
      aksi: row[2],
      detail: row[3]
    };
  });
}

// ==========================================
// FITUR STOK OPNAME (BULK & HISTORY)
// ==========================================

// 1. Simpan Stok Opname Sekaligus (Bulk)
// [UPDATE DI CODE.GS]
function simpanStokOpnameBulk(dataDraft) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetProduk = ss.getSheetByName('PRODUK');
  const sheetRiwayat = ss.getSheetByName('RIWAYAT_SO');
  
  const dataProduk = sheetProduk.getDataRange().getValues();
  const user = Session.getActiveUser().getEmail() || 'Admin';
  const idSO = 'SO-' + Date.now(); // ID Transaksi Dibuat Disini
  const tanggal = new Date();
  
  let detailHistory = [];

  // Update Stok & Catat Log
  dataDraft.forEach(item => {
    for (let i = 1; i < dataProduk.length; i++) {
      if (String(dataProduk[i][0]) === String(item.id)) {
        
        let stokIsiLama = dataProduk[i][4];
        let stokKosongLama = dataProduk[i][5];

        sheetProduk.getRange(i + 1, 5).setValue(Number(item.fisikIsi));
        sheetProduk.getRange(i + 1, 6).setValue(Number(item.fisikKosong));

        detailHistory.push({
          nama: dataProduk[i][1],
          isi_lama: stokIsiLama,
          isi_baru: item.fisikIsi,
          kosong_lama: stokKosongLama,
          kosong_baru: item.fisikKosong,
          keterangan: item.keterangan
        });
        break;
      }
    }
  });

  // Simpan ke Riwayat
  if(sheetRiwayat) {
    sheetRiwayat.appendRow([
      idSO,
      tanggal,
      user,
      dataDraft.length,
      JSON.stringify(detailHistory)
    ]);
  }

  catatLog("STOK_OPNAME", `Bulk Update ${dataDraft.length} item. ID: ${idSO}`);

  // --- TAMBAHKAN BARIS INI ---
  SpreadsheetApp.flush(); // Memaksa data tertulis segera agar bisa langsung dibaca
  // ---------------------------

  // [KUNCI UTAMA] Return Object JSON berisi ID
  return JSON.stringify({
      status: 'success',
      id: idSO, 
      message: 'Data Stok Opname Berhasil Disimpan!'
  });
}

// [PASTIKAN FUNGSI INI JUGA ADA DI PALING BAWAH CODE.GS]
function getDetailOpname(idOpname) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("RIWAYAT_SO");
  if(!sheet) return "[]";
  const data = sheet.getDataRange().getValues();
  
  for(let i=1; i<data.length; i++) {
     if(String(data[i][0]) === String(idOpname)) return data[i][4]; 
  }
  return "[]";
}

// 2. Ambil Daftar Riwayat SO
function getRiwayatSO() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('RIWAYAT_SO');
  if(!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  if(data.length <= 1) return [];
  
  // Hapus header, lalu balik urutan agar yang terbaru ada di atas
  // Kita format tanggalnya sekalian di sini biar rapi
  const timeZone = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  
  let result = data.slice(1).map(r => {
     let tgl = r[1];
     if (tgl instanceof Date) {
        tgl = Utilities.formatDate(tgl, timeZone, "dd/MM/yyyy HH:mm");
     }
     return [r[0], tgl, r[2], r[3], r[4]];
  });

  return result.reverse();
}

function generatePdfFromHtml(htmlContent, filename, logoId) { 
  if (logoId) {
    try {
      var file = DriveApp.getFileById(logoId);
      var blob = file.getBlob();
      var b64Data = Utilities.base64Encode(blob.getBytes());
      var type = blob.getContentType();
      var imgSource = "data:" + type + ";base64," + b64Data;
      
      htmlContent = htmlContent.replace('[[LOGO_IMAGE_PLACEHOLDER]]', imgSource);
      
    } catch (e) {
      htmlContent = htmlContent.replace('[[LOGO_IMAGE_PLACEHOLDER]]', ''); 
    }
  }

  var htmlBlob = Utilities.newBlob(htmlContent, 'text/html', filename + '.html');
  var tempFile = DriveApp.createFile(htmlBlob);
  var pdfBlob = tempFile.getAs('application/pdf');
  pdfBlob.setName(filename + ".pdf");
  var base64 = Utilities.base64Encode(pdfBlob.getBytes());
  tempFile.setTrashed(true);

  catatLog("EXPORT_PDF", "Mendownload file: " + filename);
   
  return base64;
}
