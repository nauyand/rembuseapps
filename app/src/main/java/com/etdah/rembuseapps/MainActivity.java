
package com.etdah.rembuseapps;

import android.content.ContentValues;
import android.net.Uri;
import android.os.Bundle;
import android.provider.MediaStore;
import android.view.View;
import android.widget.ArrayAdapter;
import android.widget.Button;
import android.widget.EditText;
import android.widget.Spinner;
import android.widget.TextView;
import android.widget.Toast;
import android.app.DatePickerDialog;

import androidx.appcompat.app.AppCompatActivity;

import java.io.InputStream;
import java.io.OutputStream;
import java.text.NumberFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;


import android.content.Intent;

import androidx.activity.result.ActivityResultLauncher;
import androidx.activity.result.contract.ActivityResultContracts;

import java.io.ByteArrayOutputStream;
import java.util.Collections;




public class MainActivity extends AppCompatActivity {
    private ActivityResultLauncher<String[]> reopenExcelPicker;
    private TextView txtJumlahFoto;
    private final List<Uri> fotoBonUris = new ArrayList<>();
    private final List<byte[]> fotoBonBytes = new ArrayList<>();
    private ActivityResultLauncher<String[]> photoPicker;

    private EditText editProject, editPIC;
    private EditText editReimburseDate, editPurchaseDate;
    private EditText editUraian, editNominal;

    private Spinner spinnerKategori;

    private TextView txtJenisTerpilih;
    private TextView txtTotalMaterial, txtTotalJasa;
    private TextView txtTotalOperasional, txtTotalSemua;
    private TextView txtDaftarTransaksi, txtStatus;

    private Button btnMaterial, btnJasa, btnOperasional;

    private String jenisTerpilih = "Material";

    private double totalMaterial = 0;
    private double totalJasa = 0;
    private double totalOperasional = 0;

    private final List<Transaksi> transaksiList = new ArrayList<>();

    private final ExecutorService executor =
            Executors.newSingleThreadExecutor();

    private final Locale localeIndonesia =
            new Locale("id", "ID");

    private final SimpleDateFormat dateFormat =
            new SimpleDateFormat("dd/MM/yyyy", Locale.US);

    private static final String FOLDER_PATH =
            "Documents/Rembuse/";

    private static final String MIME_XLSX =
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        reopenExcelPicker = registerForActivityResult(
                new ActivityResultContracts.OpenDocument(),
                uri -> {
                    if (uri == null) return;

                    try {
                        getContentResolver().takePersistableUriPermission(
                                uri,
                                Intent.FLAG_GRANT_READ_URI_PERMISSION
                        );
                    } catch (SecurityException ignored) {
                        // Tetap coba membaca file.
                    }

                    muatFileExcel(uri);
                }
        );

        findViewById(R.id.btnReopenExcel).setOnClickListener(v ->
                reopenExcelPicker.launch(new String[]{
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                })
        );


        txtJumlahFoto = findViewById(R.id.txtJumlahFoto);

        photoPicker = registerForActivityResult(
                new ActivityResultContracts.OpenMultipleDocuments(),
                uris -> {
                    if (uris == null || uris.isEmpty()) {
                        return;
                    }

                    for (Uri uri : uris) {
                        try {
                            getContentResolver().takePersistableUriPermission(
                                    uri,
                                    Intent.FLAG_GRANT_READ_URI_PERMISSION
                            );
                        } catch (SecurityException ignored) {
                            // Jika izin permanen tidak tersedia,
                            // coba akses URI saat ekspor.
                        }

                        if (!fotoBonUris.contains(uri)) {
                            fotoBonUris.add(uri);
                        }
                    }

                    txtJumlahFoto.setText(
                            "Lampiran bon: " + fotoBonUris.size() + " foto");
                }
        );

        findViewById(R.id.btnPilihFoto).setOnClickListener(v ->
                photoPicker.launch(new String[]{
                        "image/jpeg",
                        "image/png"
                })
        );

        hubungkanKomponen();
        aturPemilihTanggal();
        aturKategori();
        aturTombolJenis();

        findViewById(R.id.btnTambahTransaksi)
                .setOnClickListener(v -> tambahTransaksi());

        findViewById(R.id.btnEksporExcel)
                .setOnClickListener(v -> eksporExcel());

        perbaruiRekap();
    }

    private List<List<String>> bacaRentangExcel(
            Uri uri,
            int barisAwal,
            int barisAkhir
    ) throws Exception {

        try (InputStream input =
                     getContentResolver().openInputStream(uri)) {

            if (input == null) {
                throw new Exception("File Excel tidak dapat dibaca.");
            }

            return ExcelHelper.readRange(
                    input,
                    "Rincian",
                    barisAwal,
                    barisAkhir,
                    1,
                    5
            );
        }
    }

    private void tambahkanTransaksiDariBaris(
            List<Transaksi> hasil,
            List<List<String>> baris,
            String jenis
    ) {
        for (List<String> kolom : baris) {
            String tanggal = kolom.get(1);
            String nominalTeks = kolom.get(2);
            String kategori = kolom.get(3);
            String uraian = kolom.get(4);

            android.util.Log.d(
                    "REOPEN_DEBUG",
                    "Jenis=" + jenis
                            + " | Tanggal=" + tanggal
                            + " | Nominal mentah=" + nominalTeks
                            + " | Kategori=" + kategori
                            + " | Uraian=" + uraian
            );
            // Baris kosong tidak dianggap transaksi.
            if (tanggal.isEmpty()
                    && nominalTeks.isEmpty()
                    && kategori.isEmpty()
                    && uraian.isEmpty()) {
                continue;
            }

            // Nominal kosong atau nol tidak dimuat.
            String angka = nominalTeks.replaceAll("[^0-9-]", "");

            if (angka.isEmpty() || angka.equals("-")) {
                continue;
            }

            double nominal;

            try {
                nominal = Double.parseDouble(angka);
            } catch (NumberFormatException e) {
                continue;
            }

            if (nominal <= 0) {
                continue;
            }

            hasil.add(new Transaksi(
                    jenis,
                    tanggal,
                    kategori,
                    uraian,
                    nominal
            ));
        }
    }
    private void muatFileExcel(Uri uri) {
        findViewById(R.id.btnReopenExcel).setEnabled(false);
        txtStatus.setText("Sedang memuat data reimbursement...");

        executor.execute(() -> {
            try {
                // Baca informasi utama dari sheet Resume.
                String proyek;
                String pic;
                String tanggal;

                try (InputStream input =
                             getContentResolver().openInputStream(uri)) {
                    if (input == null) {
                        throw new Exception("File Excel tidak dapat dibuka.");
                    }

                    proyek = ExcelHelper.readCell(input, "Resume", "G10");
                }

                try (InputStream input =
                             getContentResolver().openInputStream(uri)) {
                    if (input == null) {
                        throw new Exception("File Excel tidak dapat dibuka.");
                    }

                    pic = ExcelHelper.readCell(input, "Resume", "G12");
                }

                try (InputStream input =
                             getContentResolver().openInputStream(uri)) {
                    if (input == null) {
                        throw new Exception("File Excel tidak dapat dibuka.");
                    }

                    tanggal = ExcelHelper.readCell(input, "Resume", "G11");
                }

                // Baca rincian berdasarkan kelompok baris.
                List<Transaksi> transaksiBaru = new ArrayList<>();

                tambahkanTransaksiDariBaris(
                        transaksiBaru,
                        bacaRentangExcel(uri, 9, 37),
                        "Material"
                );

                tambahkanTransaksiDariBaris(
                        transaksiBaru,
                        bacaRentangExcel(uri, 43, 49),
                        "Jasa"
                );

                tambahkanTransaksiDariBaris(
                        transaksiBaru,
                        bacaRentangExcel(uri, 55, 117),
                        "Operasional"
                );

                android.util.Log.d(
                        "REMBUSE_REKAP",
                        "Jumlah transaksi terbaca: " + transaksiBaru.size()
                );

                for (Transaksi t : transaksiBaru) {
                    android.util.Log.d(
                            "REMBUSE_REKAP",
                            "Jenis=" + t.jenis
                                    + " | Nominal=" + t.nominal
                                    + " | Kategori=" + t.kategori
                                    + " | Uraian=" + t.uraian
                    );
                }

                List<byte[]> fotoLama;

                try (InputStream input =
                             getContentResolver().openInputStream(uri)) {

                    if (input == null) {
                        throw new Exception("File Excel tidak dapat dibuka.");
                    }

                    fotoLama = ExcelHelper.readPictures(input, "Lampiran");
                }

                android.util.Log.d(
                        "REOPEN_FOTO",
                        "Jumlah foto yang berhasil dibaca: " + fotoLama.size()
                );

                runOnUiThread(() -> {
                    // Ganti data di memori dengan data file yang dibuka.
                    transaksiList.clear();
                    transaksiList.addAll(transaksiBaru);

                    editProject.setText(proyek);
                    editPIC.setText(pic);
                    editReimburseDate.setText(tanggal);

                    // Perbarui daftar transaksi dan rekap di layar.
                    perbaruiRekap();
                    fotoBonBytes.clear();
                    fotoBonBytes.addAll(fotoLama);

                    txtJumlahFoto.setText(
                            "Lampiran bon: " + fotoBonBytes.size() + " foto"
                    );
                    txtStatus.setText(
                            "File berhasil dimuat. "
                                    + transaksiList.size()
                                    + " transaksi ditemukan."
                    );

                    findViewById(R.id.btnReopenExcel).setEnabled(true);

                    Toast.makeText(
                            MainActivity.this,
                            "Berhasil memuat "
                                    + transaksiList.size()
                                    + " transaksi",
                            Toast.LENGTH_LONG
                    ).show();
                });

            } catch (Exception e) {
                android.util.Log.e(
                        "REMBUSE_EXCEL",
                        "Gagal memuat reimbursement",
                        e
                );

                runOnUiThread(() -> {
                    txtStatus.setText(
                            "Gagal memuat file: " + e.getMessage()
                    );

                    findViewById(R.id.btnReopenExcel).setEnabled(true);

                    Toast.makeText(
                            MainActivity.this,
                            "Gagal memuat data Excel. Periksa Logcat.",
                            Toast.LENGTH_LONG
                    ).show();
                });
            }
        });
    }

    private void hubungkanKomponen() {
        editProject = findViewById(R.id.editProject);
        editPIC = findViewById(R.id.editPIC);

        editReimburseDate =
                findViewById(R.id.editReimburseDate);

        editPurchaseDate =
                findViewById(R.id.editPurchaseDate);

        editUraian = findViewById(R.id.editUraian);
        editNominal = findViewById(R.id.editNominal);

        spinnerKategori = findViewById(R.id.spinnerKategori);

        txtJenisTerpilih =
                findViewById(R.id.txtJenisTerpilih);

        txtTotalMaterial =
                findViewById(R.id.txtTotalMaterial);

        txtTotalJasa =
                findViewById(R.id.txtTotalJasa);

        txtTotalOperasional =
                findViewById(R.id.txtTotalOperasional);

        txtTotalSemua =
                findViewById(R.id.txtTotalSemua);

        txtDaftarTransaksi =
                findViewById(R.id.txtDaftarTransaksi);

        txtStatus = findViewById(R.id.txtStatus);

        btnMaterial = findViewById(R.id.btnMaterial);
        btnJasa = findViewById(R.id.btnJasa);
        btnOperasional = findViewById(R.id.btnOperasional);
    }

    private void aturPemilihTanggal() {
        editReimburseDate.setOnClickListener(v ->
                tampilkanDatePicker(editReimburseDate));

        editPurchaseDate.setOnClickListener(v ->
                tampilkanDatePicker(editPurchaseDate));
    }

    private void tampilkanDatePicker(EditText target) {
        Calendar calendar = Calendar.getInstance();

        String tanggalSaatIni = target.getText().toString().trim();

        if (!tanggalSaatIni.isEmpty()) {
            try {
                Date tanggal = dateFormat.parse(tanggalSaatIni);

                if (tanggal != null) {
                    calendar.setTime(tanggal);
                }
            } catch (Exception ignored) {
                // Gunakan tanggal hari ini jika format tidak cocok.
            }
        }

        new DatePickerDialog(
                this,
                (view, year, month, dayOfMonth) -> {
                    Calendar selected = Calendar.getInstance();

                    selected.set(year, month, dayOfMonth);

                    target.setText(dateFormat.format(
                            selected.getTime()));
                },
                calendar.get(Calendar.YEAR),
                calendar.get(Calendar.MONTH),
                calendar.get(Calendar.DAY_OF_MONTH)
        ).show();
    }

    private void aturKategori() {
        btnMaterial.setOnClickListener(v -> {
            pilihJenis("Material");
            aturTombolJenis();
        });

        btnJasa.setOnClickListener(v -> {
            pilihJenis("Jasa");
            aturTombolJenis();
        });

        btnOperasional.setOnClickListener(v -> {
            pilihJenis("Operasional");
            aturTombolJenis();
        });

        perbaruiKategori();
        aturTombolJenis();
    }

    private void pilihJenis(String jenis) {
        jenisTerpilih = jenis;

        txtJenisTerpilih.setText("Jenis: " + jenis);

        perbaruiKategori();
    }

    private void perbaruiKategori() {
        List<String> kategori = new ArrayList<>();

        switch (jenisTerpilih) {
            case "Material":
                kategori.add("Utama");
                kategori.add("Pendukung");
                break;

            case "Jasa":
                kategori.add("Axius");
                kategori.add("Subcon");
                break;

            default:
                kategori.add("Transportasi");
                kategori.add("Meal");
                kategori.add("Akomodasi");
                kategori.add("ATK");
                kategori.add("Entertainment");
                kategori.add("Toolkit");
                break;
        }

        ArrayAdapter<String> adapter =
                new ArrayAdapter<>(
                        this,
                        android.R.layout.simple_spinner_item,
                        kategori
                );

        adapter.setDropDownViewResource(
                android.R.layout.simple_spinner_dropdown_item);

        spinnerKategori.setAdapter(adapter);
    }

    private void aturTombolJenis() {
        btnMaterial.setAlpha(
                "Material".equals(jenisTerpilih) ? 1.0f : 0.5f);

        btnJasa.setAlpha(
                "Jasa".equals(jenisTerpilih) ? 1.0f : 0.5f);

        btnOperasional.setAlpha(
                "Operasional".equals(jenisTerpilih) ? 1.0f : 0.5f);
    }
    private void tambahTransaksi() {
        String tanggal = editPurchaseDate.getText()
                .toString().trim();

        String uraian = editUraian.getText()
                .toString().trim();

        String nominalText = editNominal.getText()
                .toString().trim();

        if (tanggal.isEmpty()) {
            editPurchaseDate.setError(
                    "Tanggal pembelian wajib diisi");
            return;
        }

        if (uraian.isEmpty()) {
            editUraian.setError("Uraian wajib diisi");
            return;
        }

        if (nominalText.isEmpty()) {
            editNominal.setError("Nominal wajib diisi");
            return;
        }

        double nominal;

        try {
            nominal = Double.parseDouble(
                    nominalText.replace(",", "."));

            if (!Double.isFinite(nominal) || nominal <= 0) {
                throw new NumberFormatException();
            }

        } catch (NumberFormatException e) {
            editNominal.setError(
                    "Masukkan nominal lebih besar dari nol");
            return;
        }

        String kategori =
                spinnerKategori.getSelectedItem().toString();

        Transaksi transaksi = new Transaksi(
                jenisTerpilih,
                tanggal,
                kategori,
                uraian,
                nominal
        );

        transaksiList.add(transaksi);

        switch (jenisTerpilih) {
            case "Material":
                totalMaterial += nominal;
                break;

            case "Jasa":
                totalJasa += nominal;
                break;

            case "Operasional":
                totalOperasional += nominal;
                break;
        }

        editPurchaseDate.setText("");
        editUraian.setText("");
        editNominal.setText("");

        perbaruiRekap();

        Toast.makeText(
                this,
                "Transaksi berhasil ditambahkan",
                Toast.LENGTH_SHORT
        ).show();
    }

    private void perbaruiRekap() {

        // Hitung ulang total berdasarkan transaksiList terbaru
        totalMaterial = 0;
        totalJasa = 0;
        totalOperasional = 0;

        for (Transaksi t : transaksiList) {

            switch (t.jenis) {

                case "Material":
                    totalMaterial += t.nominal;
                    break;

                case "Jasa":
                    totalJasa += t.nominal;
                    break;

                case "Operasional":
                    totalOperasional += t.nominal;
                    break;
            }
        }

        NumberFormat rupiah =
                NumberFormat.getCurrencyInstance(localeIndonesia);

        rupiah.setMaximumFractionDigits(0);

        txtTotalMaterial.setText(
                "Material: " + rupiah.format(totalMaterial));

        txtTotalJasa.setText(
                "Jasa: " + rupiah.format(totalJasa));

        txtTotalOperasional.setText(
                "Operasional: " + rupiah.format(totalOperasional));

        double total =
                totalMaterial + totalJasa + totalOperasional;

        txtTotalSemua.setText(
                "TOTAL: " + rupiah.format(total));

        if (transaksiList.isEmpty()) {
            txtDaftarTransaksi.setText("Belum ada transaksi.");
            return;
        }

        StringBuilder daftar = new StringBuilder();

        for (int i = 0; i < transaksiList.size(); i++) {

            Transaksi t = transaksiList.get(i);

            daftar.append(i + 1)
                    .append(". ")
                    .append(t.jenis)
                    .append(" - ")
                    .append(t.kategori)
                    .append("\n")
                    .append(t.tanggal)
                    .append(" | ")
                    .append(t.uraian)
                    .append("\n")
                    .append(rupiah.format(t.nominal))
                    .append("\n\n");
        }

        txtDaftarTransaksi.setText(daftar.toString());
    }

    private void eksporExcel() {
        String proyek =
                editProject.getText().toString().trim();

        String pic =
                editPIC.getText().toString().trim();

        String tanggalReimburse =
                editReimburseDate.getText().toString().trim();

        if (proyek.isEmpty()) {
            editProject.setError("Nama proyek wajib diisi");
            return;
        }

        if (pic.isEmpty()) {
            editPIC.setError("PIC wajib diisi");
            return;
        }

        if (tanggalReimburse.isEmpty()) {
            editReimburseDate.setError(
                    "Tanggal reimburse wajib diisi");
            return;
        }

        // Snapshot agar data yang diekspor konsisten.
        final String proyekFinal = proyek;
        final String picFinal = pic;
        final String tanggalFinal = tanggalReimburse;
        final List<Transaksi> transaksiFinal =
                new ArrayList<>(transaksiList);

        findViewById(R.id.btnEksporExcel).setEnabled(false);
        txtStatus.setText("Sedang membuat file Excel...");

        final List<byte[]> fotoLamaFinal =
                new ArrayList<>(fotoBonBytes);

        final List<Uri> fotoBaruFinal =
                new ArrayList<>(fotoBonUris);

        executor.execute(() -> {
            Uri outputUri = null;

            try {
                // Foto dari file reimbursement yang dibuka sebelumnya.
                List<byte[]> fotoBytes =
                        new ArrayList<>(fotoLamaFinal);

                // Foto tambahan yang baru dipilih dari galeri.
                for (Uri fotoUri : fotoBaruFinal) {
                    fotoBytes.add(bacaFoto(fotoUri));
                }

                // Lanjutkan kode ekspor yang sudah ada di sini.

                String timestamp = new SimpleDateFormat(
                        "yyyyMMdd_HHmmss",
                        Locale.US
                ).format(new Date());
                String projectFileName = proyekFinal
                        .replaceAll("[\\\\/:*?\"<>|]", "_")
                        .replaceAll("\\s+", "_")
                        .trim();

                String timestampSingkat = new SimpleDateFormat(
                        "yyMMdd",
                        Locale.US
                ).format(new Date());

                String fileName = "Reimbursement_"
                        + projectFileName
                        + "_"
                        + timestampSingkat
                        + ".xlsx";

                Map<String, Map<String, Object>> allData =
                        new LinkedHashMap<>();

                Map<String, Object> resume =
                        new LinkedHashMap<>();

                resume.put("G10", proyekFinal);
                resume.put("G11", tanggalFinal);
                resume.put("G12", picFinal);

                resume.put("H15", totalDari(
                        transaksiFinal, "Material"));

                resume.put("H16", totalDari(
                        transaksiFinal, "Jasa"));

                resume.put("H17", totalDari(
                        transaksiFinal, "Operasional"));

                double totalSemua =
                        totalDari(transaksiFinal, "Material")
                                + totalDari(transaksiFinal, "Jasa")
                                + totalDari(transaksiFinal, "Operasional");

                resume.put("H18", totalSemua);
                resume.put("H20", 3000000.0);
                resume.put("H22", 3000000.0 - totalSemua);

                allData.put("Resume", resume);

                Map<String, Object> rincian =
                        buatDataRincian(transaksiFinal);

                allData.put("Rincian", rincian);

                ContentValues values = new ContentValues();

                values.put(
                        MediaStore.MediaColumns.DISPLAY_NAME,
                        fileName);

                values.put(
                        MediaStore.MediaColumns.MIME_TYPE,
                        MIME_XLSX);

                values.put(
                        MediaStore.MediaColumns.RELATIVE_PATH,
                        FOLDER_PATH);

                values.put(
                        MediaStore.MediaColumns.IS_PENDING,
                        1);

                Uri collection =
                        MediaStore.Files.getContentUri(
                                MediaStore.VOLUME_EXTERNAL_PRIMARY);

                outputUri = getContentResolver().insert(
                        collection, values);

                if (outputUri == null) {
                    throw new Exception(
                            "Gagal membuat file Excel.");
                }

                try (InputStream template =
                             getAssets().open("TEMPLATE.xlsx");
                     OutputStream output =
                             getContentResolver()
                                     .openOutputStream(outputUri)) {

                    if (output == null) {
                        throw new Exception(
                                "Gagal membuka file output.");
                    }


                    ExcelHelper.writeWorkbook(
                            template,
                            output,
                            allData,
                            fotoBytes
                    );

                }

                // Pastikan hasil bisa dibaca sebelum dipublikasikan.
                try (InputStream verify =
                             getContentResolver()
                                     .openInputStream(outputUri)) {

                    if (verify == null) {
                        throw new Exception(
                                "Gagal membaca hasil Excel.");
                    }

                    String hasil = ExcelHelper.readCell(
                            verify, "Resume", "G10");

                    if (!proyekFinal.equals(hasil)) {
                        throw new Exception(
                                "Verifikasi Resume!G10 gagal: "
                                        + hasil);
                    }
                }



                ContentValues published = new ContentValues();

                published.put(
                        MediaStore.MediaColumns.IS_PENDING, 0);

                getContentResolver().update(
                        outputUri, published, null, null);

                runOnUiThread(() -> {
                    txtStatus.setText(
                            "Ekspor berhasil!\n"
                                    + "Documents/Rembuse/"
                                    + fileName);

                    Toast.makeText(
                            this,
                            "Excel berhasil disimpan",
                            Toast.LENGTH_LONG
                    ).show();

                    findViewById(R.id.btnEksporExcel)
                            .setEnabled(true);
                });

            } catch (Exception e) {
                android.util.Log.e(
                        "REMBUSE_EXCEL",
                        "Gagal ekspor Excel",
                        e);

                if (outputUri != null) {
                    getContentResolver().delete(
                            outputUri, null, null);
                }

                runOnUiThread(() -> {
                    txtStatus.setText(
                            "Ekspor gagal: " + e.getMessage());

                    Toast.makeText(
                            this,
                            "Ekspor gagal. Periksa Logcat.",
                            Toast.LENGTH_LONG
                    ).show();

                    findViewById(R.id.btnEksporExcel)
                            .setEnabled(true);
                });
            }
        });
    }

    private double totalDari(
            List<Transaksi> data,
            String jenis
    ) {
        double total = 0;

        for (Transaksi t : data) {
            if (t.jenis.equals(jenis)) {
                total += t.nominal;
            }
        }

        return total;
    }

    private Map<String, Object> buatDataRincian(
            List<Transaksi> data
    ) throws Exception {

        Map<String, Object> cells =
                new LinkedHashMap<>();

        // Pemetaan awal untuk diperiksa pada file hasil.
        // A = nomor, B = tanggal, C = nominal,
        // D = kategori, E = uraian.
        //
        // Rentang awal:
        // Material     : baris 9-37, total C38
        // Jasa         : baris 43-49, total C50
        // Operasional  : baris 60-117, total C118
        //
        // Operasional dimulai dari baris 60 untuk
        // menghindari menimpa sel total Jasa di C50.

        int materialRow = 9;
        int jasaRow = 43;
        int operasionalRow = 55;
        int nomorMaterial = 1;
        int nomorJasa = 1;
        int nomorOperasional = 1;
        double subtotalMaterial = 0;
        double subtotalJasa = 0;
        double subtotalOperasional = 0;

        for (Transaksi t : data) {
            int row;

            switch (t.jenis) {
                case "Material":
                    if (materialRow > 37) {
                        throw new Exception("Kapasitas Material melebihi baris 9-37.");
                    }

                    row = materialRow++;
                    cells.put("A" + row, nomorMaterial++);
                    subtotalMaterial += t.nominal;
                    break;

                case "Jasa":
                    if (jasaRow > 49) {
                        throw new Exception("Kapasitas Jasa melebihi baris 43-49.");
                    }

                    row = jasaRow++;
                    cells.put("A" + row, nomorJasa++);
                    subtotalJasa += t.nominal;
                    break;

                default:
                    if (operasionalRow > 117) {
                        throw new Exception("Kapasitas Operasional melebihi baris 55-117.");
                    }

                    row = operasionalRow++;
                    cells.put("A" + row, nomorOperasional++);
                    subtotalOperasional += t.nominal;
                    break;
            }

            cells.put("B" + row, t.tanggal);
            cells.put("C" + row, t.nominal);
            cells.put("D" + row, t.kategori);
            cells.put("E" + row, t.uraian);
        }

        cells.put("C38", subtotalMaterial);
        cells.put("C50", subtotalJasa);
        cells.put("C118", subtotalOperasional);

        return cells;
    }

    private byte[] bacaFoto(Uri uri) throws Exception {
        try (InputStream input =
                     getContentResolver().openInputStream(uri);
             ByteArrayOutputStream output =
                     new ByteArrayOutputStream()) {

            if (input == null) {
                throw new Exception("Foto tidak dapat dibaca: " + uri);
            }

            byte[] buffer = new byte[8192];
            int jumlah;

            while ((jumlah = input.read(buffer)) != -1) {
                output.write(buffer, 0, jumlah);
            }

            return output.toByteArray();
        }
    }

    private int nomorUntuk(
            Transaksi transaksi,
            List<Transaksi> data
    ) {
        return data.indexOf(transaksi) + 1;
    }

    private static class Transaksi {
        final String jenis;
        final String tanggal;
        final String kategori;
        final String uraian;
        final double nominal;

        Transaksi(
                String jenis,
                String tanggal,
                String kategori,
                String uraian,
                double nominal
        ) {
            this.jenis = jenis;
            this.tanggal = tanggal;
            this.kategori = kategori;
            this.uraian = uraian;
            this.nominal = nominal;
        }
    }

    @Override
    protected void onDestroy() {
        executor.shutdown();
        super.onDestroy();
    }
}
