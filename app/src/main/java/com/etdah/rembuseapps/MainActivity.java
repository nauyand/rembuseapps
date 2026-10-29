package com.etdah.rembuseapps;

import android.app.DatePickerDialog;
import android.content.ClipData;
import android.content.ClipboardManager;
import android.content.ContentValues;
import android.content.Intent;
import android.graphics.Color;
import android.net.Uri;
import android.os.Bundle;
import android.provider.MediaStore;
import android.text.Editable;
import android.text.TextWatcher;
import android.util.Log;
import android.util.TypedValue;
import android.view.View;
import android.widget.ArrayAdapter;
import android.widget.Button;
import android.widget.EditText;
import android.widget.LinearLayout;
import android.widget.Spinner;
import android.widget.TextView;
import android.widget.Toast;

import androidx.activity.result.ActivityResultLauncher;
import androidx.activity.result.contract.ActivityResultContracts;
import androidx.appcompat.app.AlertDialog;
import androidx.appcompat.app.AppCompatActivity;

import org.json.JSONArray;
import org.json.JSONObject;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.nio.charset.StandardCharsets;
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

public class MainActivity extends AppCompatActivity {

    private static final String TAG = "REMBUSE";

    private static final String FOLDER_PATH = "Documents/Rembuse/";

    private static final String MIME_XLSX =
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

    /** Nilai kas awal jika belum pernah diisi. */
    private static final double KAS_DEFAULT = 3000000.0;

    /** File draf agar data tidak hilang saat layar diputar / app ditutup. */
    private static final String DRAFT_FILE = "draft.json";

    /** Folder foto bon yang berasal dari file Excel yang dibuka ulang. */
    private static final String DRAFT_FOTO_DIR = "draft_foto";

    private static final String KEY_EDIT_INDEX = "editIndex";

    /**
     * Executor dibuat static supaya tetap satu antrean walau activity
     * dibuat ulang (misalnya saat layar diputar).
     */
    private static final ExecutorService executor =
            Executors.newSingleThreadExecutor();

    private final Locale localeIndonesia = new Locale("id", "ID");

    private final SimpleDateFormat dateFormat =
            new SimpleDateFormat("dd/MM/yyyy", Locale.US);

    // ---------- Komponen layar ----------
    private EditText editProject, editPIC, editKas;
    private EditText editReimburseDate, editPurchaseDate;
    private EditText editUraian, editNominal;
    private Spinner spinnerKategori;
    private TextView txtJenisTerpilih;
    private TextView txtTotalMaterial, txtTotalJasa;
    private TextView txtTotalOperasional, txtTotalSemua, txtSaldo;
    private TextView txtDaftarTransaksi, txtStatus, txtJumlahFoto;
    private LinearLayout listTransaksi;
    private Button btnMaterial, btnJasa, btnOperasional;
    private Button btnTambahTransaksi, btnBatalEdit;

    // ---------- Data ----------
    private String jenisTerpilih = "Material";
    private final List<Transaksi> transaksiList = new ArrayList<>();

    /** Foto bon baru yang dipilih dari galeri. */
    private final List<Uri> fotoBonUris = new ArrayList<>();

    /** Foto bon dari file Excel yang dibuka ulang (disimpan di memori internal). */
    private final List<File> fotoLamaFiles = new ArrayList<>();

    /** Posisi transaksi yang sedang diedit, -1 jika sedang menambah baru. */
    private int editIndex = -1;

    private ActivityResultLauncher<String[]> reopenExcelPicker;
    private ActivityResultLauncher<String[]> photoPicker;

    // =====================================================================
    // Siklus hidup activity
    // =====================================================================

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        hubungkanKomponen();
        daftarkanPemilihFile();
        aturPemilihTanggal();
        aturKategori();

        btnTambahTransaksi.setOnClickListener(v -> simpanTransaksi());
        btnBatalEdit.setOnClickListener(v -> keluarModeEdit(true));

        findViewById(R.id.btnEksporExcel)
                .setOnClickListener(v -> eksporExcel());

        findViewById(R.id.btnHapusFoto)
                .setOnClickListener(v -> konfirmasiHapusFoto());

        findViewById(R.id.btnMulaiBaru)
                .setOnClickListener(v -> konfirmasiMulaiBaru());

        // Pulihkan data terakhir (setelah layar diputar / app ditutup).
        muatDraf();

        // Rekap ikut berubah setiap kali jumlah kas diketik.
        editKas.addTextChangedListener(new TextWatcher() {
            @Override
            public void beforeTextChanged(CharSequence s, int a, int b, int c) {
            }

            @Override
            public void onTextChanged(CharSequence s, int a, int b, int c) {
            }

            @Override
            public void afterTextChanged(Editable s) {
                perbaruiRekap();
            }
        });

        // Jika layar diputar saat sedang mengedit, lanjutkan mode edit.
        // Isian form sendiri dipulihkan otomatis oleh Android.
        if (savedInstanceState != null) {
            int index = savedInstanceState.getInt(KEY_EDIT_INDEX, -1);
            if (index >= 0 && index < transaksiList.size()) {
                editIndex = index;
                btnTambahTransaksi.setText("✓  Simpan Perubahan");
                btnBatalEdit.setVisibility(View.VISIBLE);
            }
        }

        perbaruiTombolJenis();
        perbaruiJumlahFoto();
        perbaruiRekap();
    }

    @Override
    protected void onSaveInstanceState(Bundle outState) {
        super.onSaveInstanceState(outState);
        outState.putInt(KEY_EDIT_INDEX, editIndex);
    }

    @Override
    protected void onPause() {
        super.onPause();
        simpanDraf();
    }

    // =====================================================================
    // Inisialisasi
    // =====================================================================

    private void hubungkanKomponen() {
        editProject = findViewById(R.id.editProject);
        editPIC = findViewById(R.id.editPIC);
        editKas = findViewById(R.id.editKas);
        editReimburseDate = findViewById(R.id.editReimburseDate);
        editPurchaseDate = findViewById(R.id.editPurchaseDate);
        editUraian = findViewById(R.id.editUraian);
        editNominal = findViewById(R.id.editNominal);

        spinnerKategori = findViewById(R.id.spinnerKategori);

        txtJenisTerpilih = findViewById(R.id.txtJenisTerpilih);
        txtTotalMaterial = findViewById(R.id.txtTotalMaterial);
        txtTotalJasa = findViewById(R.id.txtTotalJasa);
        txtTotalOperasional = findViewById(R.id.txtTotalOperasional);
        txtTotalSemua = findViewById(R.id.txtTotalSemua);
        txtSaldo = findViewById(R.id.txtSaldo);
        txtDaftarTransaksi = findViewById(R.id.txtDaftarTransaksi);
        txtStatus = findViewById(R.id.txtStatus);
        txtJumlahFoto = findViewById(R.id.txtJumlahFoto);
        listTransaksi = findViewById(R.id.listTransaksi);

        btnMaterial = findViewById(R.id.btnMaterial);
        btnJasa = findViewById(R.id.btnJasa);
        btnOperasional = findViewById(R.id.btnOperasional);
        btnTambahTransaksi = findViewById(R.id.btnTambahTransaksi);
        btnBatalEdit = findViewById(R.id.btnBatalEdit);
    }

    private void daftarkanPemilihFile() {
        reopenExcelPicker = registerForActivityResult(
                new ActivityResultContracts.OpenDocument(),
                uri -> {
                    if (uri == null) return;

                    try {
                        getContentResolver().takePersistableUriPermission(
                                uri, Intent.FLAG_GRANT_READ_URI_PERMISSION);
                    } catch (SecurityException ignored) {
                        // Tetap coba membaca file.
                    }

                    muatFileExcel(uri);
                }
        );

        findViewById(R.id.btnReopenExcel).setOnClickListener(v ->
                reopenExcelPicker.launch(new String[]{MIME_XLSX}));

        photoPicker = registerForActivityResult(
                new ActivityResultContracts.OpenMultipleDocuments(),
                uris -> {
                    if (uris == null || uris.isEmpty()) return;

                    for (Uri uri : uris) {
                        try {
                            // Izin permanen agar foto masih bisa dibaca
                            // setelah aplikasi ditutup lalu dibuka lagi.
                            getContentResolver().takePersistableUriPermission(
                                    uri, Intent.FLAG_GRANT_READ_URI_PERMISSION);
                        } catch (SecurityException ignored) {
                            // Coba akses URI saat ekspor.
                        }

                        if (!fotoBonUris.contains(uri)) {
                            fotoBonUris.add(uri);
                        }
                    }

                    perbaruiJumlahFoto();
                    simpanDraf();
                }
        );

        findViewById(R.id.btnPilihFoto).setOnClickListener(v ->
                photoPicker.launch(new String[]{"image/jpeg", "image/png"}));
    }

    // =====================================================================
    // Tanggal
    // =====================================================================

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
                if (tanggal != null) calendar.setTime(tanggal);
            } catch (Exception ignored) {
                // Gunakan tanggal hari ini jika format tidak cocok.
            }
        }

        new DatePickerDialog(
                this,
                (view, year, month, dayOfMonth) -> {
                    Calendar selected = Calendar.getInstance();
                    selected.set(year, month, dayOfMonth);
                    target.setText(dateFormat.format(selected.getTime()));
                    target.setError(null);
                },
                calendar.get(Calendar.YEAR),
                calendar.get(Calendar.MONTH),
                calendar.get(Calendar.DAY_OF_MONTH)
        ).show();
    }

    // =====================================================================
    // Jenis & kategori
    // =====================================================================

    private void aturKategori() {
        btnMaterial.setOnClickListener(v -> pilihJenis("Material"));
        btnJasa.setOnClickListener(v -> pilihJenis("Jasa"));
        btnOperasional.setOnClickListener(v -> pilihJenis("Operasional"));

        perbaruiKategori(null);
    }

    private void pilihJenis(String jenis) {
        jenisTerpilih = jenis;
        perbaruiKategori(null);
        perbaruiTombolJenis();
    }

    private List<String> daftarKategori(String jenis) {
        List<String> kategori = new ArrayList<>();

        switch (jenis) {
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

        return kategori;
    }

    /**
     * Mengisi ulang pilihan kategori sesuai jenis.
     *
     * @param pilih kategori yang ingin dipilih (boleh null). Jika kategori
     *              ini tidak ada di daftar (misalnya dari file Excel lama),
     *              kategori tersebut ditambahkan supaya tidak hilang.
     */
    private void perbaruiKategori(String pilih) {
        List<String> kategori = daftarKategori(jenisTerpilih);

        if (pilih != null && !pilih.isEmpty() && !kategori.contains(pilih)) {
            kategori.add(pilih);
        }

        ArrayAdapter<String> adapter = new ArrayAdapter<>(
                this, android.R.layout.simple_spinner_item, kategori);

        adapter.setDropDownViewResource(
                android.R.layout.simple_spinner_dropdown_item);

        spinnerKategori.setAdapter(adapter);

        if (pilih != null) {
            int posisi = kategori.indexOf(pilih);
            if (posisi >= 0) spinnerKategori.setSelection(posisi);
        }
    }

    private void perbaruiTombolJenis() {
        txtJenisTerpilih.setText("Jenis: " + jenisTerpilih);

        btnMaterial.setAlpha("Material".equals(jenisTerpilih) ? 1.0f : 0.5f);
        btnJasa.setAlpha("Jasa".equals(jenisTerpilih) ? 1.0f : 0.5f);
        btnOperasional.setAlpha("Operasional".equals(jenisTerpilih) ? 1.0f : 0.5f);
    }

    // =====================================================================
    // FITUR 1: Tambah, edit, dan hapus transaksi
    // =====================================================================

    /** Dipanggil tombol Tambah / Simpan Perubahan. */
    private void simpanTransaksi() {
        String tanggal = editPurchaseDate.getText().toString().trim();
        String uraian = editUraian.getText().toString().trim();
        String nominalText = editNominal.getText().toString().trim();

        if (tanggal.isEmpty()) {
            editPurchaseDate.setError("Tanggal pembelian wajib diisi");
            return;
        }

        if (uraian.isEmpty()) {
            editUraian.setError("Uraian wajib diisi");
            return;
        }

        double nominal = parseAngka(nominalText);

        if (!(nominal > 0)) {
            editNominal.setError("Masukkan nominal lebih besar dari nol");
            return;
        }

        Object item = spinnerKategori.getSelectedItem();
        String kategori = item == null ? "" : item.toString();

        Transaksi transaksi = new Transaksi(
                jenisTerpilih, tanggal, kategori, uraian, nominal);

        boolean modeEdit =
                editIndex >= 0 && editIndex < transaksiList.size();

        if (modeEdit) {
            transaksiList.set(editIndex, transaksi);
        } else {
            transaksiList.add(transaksi);
        }

        keluarModeEdit(false);
        kosongkanFormTransaksi();
        perbaruiRekap();
        simpanDraf();

        Toast.makeText(
                this,
                modeEdit ? "Transaksi diperbarui" : "Transaksi berhasil ditambahkan",
                Toast.LENGTH_SHORT
        ).show();
    }

    private void kosongkanFormTransaksi() {
        editPurchaseDate.setText("");
        editUraian.setText("");
        editNominal.setText("");
        editPurchaseDate.setError(null);
        editUraian.setError(null);
        editNominal.setError(null);
    }

    /** Menu yang muncul saat sebuah transaksi diketuk. */
    private void tampilkanMenuTransaksi(int index) {
        if (index < 0 || index >= transaksiList.size()) return;

        Transaksi t = transaksiList.get(index);

        new AlertDialog.Builder(this)
                .setTitle((index + 1) + ". " + t.uraian)
                .setItems(new String[]{"Edit", "Hapus"}, (dialog, which) -> {
                    if (which == 0) {
                        masukModeEdit(index);
                    } else {
                        konfirmasiHapusTransaksi(index);
                    }
                })
                .setNegativeButton("Tutup", null)
                .show();
    }

    private void masukModeEdit(int index) {
        Transaksi t = transaksiList.get(index);
        editIndex = index;

        jenisTerpilih = t.jenis;
        perbaruiKategori(t.kategori);
        perbaruiTombolJenis();

        editPurchaseDate.setText(t.tanggal);
        editUraian.setText(t.uraian);
        editNominal.setText(angkaUntukInput(t.nominal));

        btnTambahTransaksi.setText("✓  Simpan Perubahan");
        btnBatalEdit.setVisibility(View.VISIBLE);
        txtStatus.setText("Mengedit transaksi no. " + (index + 1));

        // Fokus ke form agar layar otomatis menggulir ke atas.
        editUraian.requestFocus();
        editUraian.setSelection(editUraian.getText().length());

        perbaruiRekap();
    }

    /** @param kosongkanForm true jika isian form ikut dibersihkan. */
    private void keluarModeEdit(boolean kosongkanForm) {
        boolean tadinyaEdit = editIndex >= 0;
        editIndex = -1;

        btnTambahTransaksi.setText("＋  Tambah Transaksi");
        btnBatalEdit.setVisibility(View.GONE);

        if (kosongkanForm && tadinyaEdit) {
            kosongkanFormTransaksi();
            txtStatus.setText("Edit dibatalkan.");
        }

        perbaruiRekap();
    }

    private void konfirmasiHapusTransaksi(int index) {
        Transaksi t = transaksiList.get(index);

        new AlertDialog.Builder(this)
                .setTitle("Hapus transaksi?")
                .setMessage(t.jenis + " - " + t.kategori + "\n"
                        + t.uraian + "\n" + formatRupiah(t.nominal))
                .setPositiveButton("Hapus", (d, w) -> {
                    if (index >= transaksiList.size()) return;

                    transaksiList.remove(index);

                    // Sesuaikan posisi transaksi yang sedang diedit.
                    if (editIndex == index) {
                        keluarModeEdit(true);
                    } else if (editIndex > index) {
                        editIndex--;
                    }

                    perbaruiRekap();
                    simpanDraf();

                    Toast.makeText(this, "Transaksi dihapus",
                            Toast.LENGTH_SHORT).show();
                })
                .setNegativeButton("Batal", null)
                .show();
    }

    // =====================================================================
    // Rekap & daftar transaksi
    // =====================================================================

    private void perbaruiRekap() {
        double totalMaterial = totalDari(transaksiList, "Material");
        double totalJasa = totalDari(transaksiList, "Jasa");
        double totalOperasional = totalDari(transaksiList, "Operasional");
        double total = totalMaterial + totalJasa + totalOperasional;

        txtTotalMaterial.setText("Material: " + formatRupiah(totalMaterial));
        txtTotalJasa.setText("Jasa: " + formatRupiah(totalJasa));
        txtTotalOperasional.setText("Operasional: " + formatRupiah(totalOperasional));
        txtTotalSemua.setText("TOTAL: " + formatRupiah(total));

        // FITUR 2: kas & saldo
        double kas = ambilKas();
        double saldo = kas - total;

        txtSaldo.setText("Kas: " + formatRupiah(kas)
                + "\nSaldo: " + formatRupiah(saldo)
                + (saldo < 0 ? "  (kurang)" : ""));
        txtSaldo.setTextColor(saldo < 0
                ? Color.parseColor("#B3261E")
                : Color.parseColor("#6741A5"));

        tampilkanDaftarTransaksi();
    }

    private void tampilkanDaftarTransaksi() {
        // Anak pertama adalah txtDaftarTransaksi (judul/keterangan),
        // sisanya baris transaksi yang dibuat ulang setiap kali.
        int jumlahAnak = listTransaksi.getChildCount();
        if (jumlahAnak > 1) {
            listTransaksi.removeViews(1, jumlahAnak - 1);
        }

        if (transaksiList.isEmpty()) {
            txtDaftarTransaksi.setText("Belum ada transaksi.");
            return;
        }

        txtDaftarTransaksi.setText(
                transaksiList.size() + " transaksi  •  ketuk untuk edit/hapus");

        TypedValue ripple = new TypedValue();
        getTheme().resolveAttribute(
                android.R.attr.selectableItemBackground, ripple, true);

        int padding = dp(8);

        for (int i = 0; i < transaksiList.size(); i++) {
            Transaksi t = transaksiList.get(i);
            final int index = i;

            View garis = new View(this);
            garis.setBackgroundColor(Color.parseColor("#E8E2EF"));
            listTransaksi.addView(garis, new LinearLayout.LayoutParams(
                    LinearLayout.LayoutParams.MATCH_PARENT, dp(1)));

            TextView baris = new TextView(this);
            baris.setText((i + 1) + ". " + t.jenis + " - " + t.kategori + "\n"
                    + t.tanggal + " | " + t.uraian + "\n"
                    + formatRupiah(t.nominal));
            baris.setTextColor(Color.parseColor("#6640A4"));
            baris.setTextSize(TypedValue.COMPLEX_UNIT_SP, 13);
            baris.setLineSpacing(dp(3), 1f);
            baris.setPadding(padding, padding, padding, padding);
            baris.setBackgroundResource(ripple.resourceId);
            baris.setClickable(true);
            baris.setFocusable(true);
            baris.setOnClickListener(v -> tampilkanMenuTransaksi(index));

            // Tandai baris yang sedang diedit.
            if (i == editIndex) {
                baris.setBackgroundColor(Color.parseColor("#F0EAFB"));
                baris.setTypeface(baris.getTypeface(), android.graphics.Typeface.BOLD);
            }

            listTransaksi.addView(baris, new LinearLayout.LayoutParams(
                    LinearLayout.LayoutParams.MATCH_PARENT,
                    LinearLayout.LayoutParams.WRAP_CONTENT));
        }
    }

    // =====================================================================
    // FITUR 4: Foto bon (tanpa dobel)
    // =====================================================================

    private int jumlahFoto() {
        return fotoLamaFiles.size() + fotoBonUris.size();
    }

    private void perbaruiJumlahFoto() {
        txtJumlahFoto.setText("Lampiran bon: " + jumlahFoto() + " foto");
    }

    private void konfirmasiHapusFoto() {
        if (jumlahFoto() == 0) {
            Toast.makeText(this, "Belum ada foto", Toast.LENGTH_SHORT).show();
            return;
        }

        new AlertDialog.Builder(this)
                .setTitle("Hapus semua foto bon?")
                .setMessage(jumlahFoto() + " foto akan dilepas dari reimbursement ini. "
                        + "Foto asli di galeri tidak ikut terhapus.")
                .setPositiveButton("Hapus", (d, w) -> {
                    hapusSemuaFoto();
                    simpanDraf();
                })
                .setNegativeButton("Batal", null)
                .show();
    }

    private void hapusSemuaFoto() {
        fotoBonUris.clear();
        fotoLamaFiles.clear();
        hapusIsiFolder(folderFotoDraf());
        perbaruiJumlahFoto();
    }

    // =====================================================================
    // Buka ulang file Excel
    // =====================================================================

    private void muatFileExcel(Uri uri) {
        findViewById(R.id.btnReopenExcel).setEnabled(false);
        txtStatus.setText("Sedang memuat data reimbursement...");

        executor.execute(() -> {
            File salinan = new File(getCacheDir(), "reopen_tmp.xlsx");

            try {
                // Salin dulu ke file sementara, lalu buka SEKALI dari disk.
                // Jauh lebih hemat memori daripada membuka stream berkali-kali.
                try (InputStream input = getContentResolver().openInputStream(uri)) {
                    if (input == null) {
                        throw new Exception("File Excel tidak dapat dibuka.");
                    }
                    salinKeFile(input, salinan);
                }

                ExcelHelper.ReimburseData data =
                        ExcelHelper.readReimbursement(salinan);

                List<Transaksi> transaksiBaru = new ArrayList<>();
                tambahkanTransaksiDariBaris(transaksiBaru, data.material, "Material");
                tambahkanTransaksiDariBaris(transaksiBaru, data.jasa, "Jasa");
                tambahkanTransaksiDariBaris(transaksiBaru, data.operasional, "Operasional");

                // Simpan foto lama ke memori internal (bukan di RAM),
                // menggantikan foto dari file sebelumnya.
                File folder = folderFotoDraf();
                hapusIsiFolder(folder);

                List<File> fileFoto = new ArrayList<>();
                long waktu = System.currentTimeMillis();

                for (int i = 0; i < data.photos.size(); i++) {
                    File f = new File(folder, "foto_" + waktu + "_" + i + ".bin");
                    tulisFile(f, data.photos.get(i));
                    fileFoto.add(f);
                }

                // Lepas referensi foto agar memori bisa dibebaskan.
                data.photos.clear();

                final Map<String, String> resume = data.resume;
                final double kas = parseAngka(resume.get("H20"));

                runOnUiThread(() -> {
                    if (isFinishing() || isDestroyed()) return;

                    try {
                        keluarModeEdit(true);

                        transaksiList.clear();
                        transaksiList.addAll(transaksiBaru);

                        editProject.setText(teksAman(resume.get("G10")));
                        editReimburseDate.setText(teksAman(resume.get("G11")));
                        editPIC.setText(teksAman(resume.get("G12")));
                        editKas.setText(kas > 0 ? angkaUntukInput(kas) : "");

                        // FITUR 4: foto galeri sebelumnya dilepas supaya
                        // tidak ikut terekspor dua kali.
                        fotoBonUris.clear();
                        fotoLamaFiles.clear();
                        fotoLamaFiles.addAll(fileFoto);

                        perbaruiJumlahFoto();
                        perbaruiRekap();
                        simpanDraf();

                        txtStatus.setText("File berhasil dimuat. "
                                + transaksiList.size() + " transaksi, "
                                + fileFoto.size() + " foto ditemukan.");

                        Toast.makeText(MainActivity.this,
                                "Berhasil memuat " + transaksiList.size() + " transaksi",
                                Toast.LENGTH_LONG).show();
                    } catch (Throwable t) {
                        tampilkanError("Gagal menampilkan data", t);
                    } finally {
                        findViewById(R.id.btnReopenExcel).setEnabled(true);
                    }
                });

            } catch (Throwable t) {
                // Throwable (bukan hanya Exception) supaya kehabisan
                // memori pun tidak membuat aplikasi force close.
                Log.e(TAG, "Gagal memuat reimbursement", t);

                runOnUiThread(() -> {
                    if (isFinishing() || isDestroyed()) return;

                    findViewById(R.id.btnReopenExcel).setEnabled(true);
                    tampilkanError(
                            t instanceof OutOfMemoryError
                                    ? "Memori tidak cukup untuk membuka file ini"
                                    : "Gagal memuat file",
                            t);
                });
            } finally {
                salinan.delete();
            }
        });
    }

    private void tambahkanTransaksiDariBaris(
            List<Transaksi> hasil, List<List<String>> baris, String jenis) {

        for (List<String> kolom : baris) {
            String tanggal = kolom.get(1);
            String nominalTeks = kolom.get(2);
            String kategori = kolom.get(3);
            String uraian = kolom.get(4);

            // Baris kosong tidak dianggap transaksi.
            if (tanggal.isEmpty() && nominalTeks.isEmpty()
                    && kategori.isEmpty() && uraian.isEmpty()) {
                continue;
            }

            double nominal = parseAngka(nominalTeks);

            if (!(nominal > 0)) continue;

            hasil.add(new Transaksi(jenis, tanggal, kategori, uraian, nominal));
        }
    }

    // =====================================================================
    // Ekspor Excel
    // =====================================================================

    private void eksporExcel() {
        String proyek = editProject.getText().toString().trim();
        String pic = editPIC.getText().toString().trim();
        String tanggalReimburse = editReimburseDate.getText().toString().trim();

        if (proyek.isEmpty()) {
            editProject.setError("Nama proyek wajib diisi");
            return;
        }

        if (pic.isEmpty()) {
            editPIC.setError("PIC wajib diisi");
            return;
        }

        if (tanggalReimburse.isEmpty()) {
            editReimburseDate.setError("Tanggal reimburse wajib diisi");
            return;
        }

        if (editIndex >= 0) {
            Toast.makeText(this,
                    "Simpan atau batalkan edit transaksi terlebih dahulu",
                    Toast.LENGTH_LONG).show();
            return;
        }

        simpanDraf();

        // Snapshot agar data yang diekspor konsisten.
        final String proyekFinal = proyek;
        final String picFinal = pic;
        final String tanggalFinal = tanggalReimburse;
        final double kasFinal = ambilKas();
        final List<Transaksi> transaksiFinal = new ArrayList<>(transaksiList);
        final List<File> fotoLamaFinal = new ArrayList<>(fotoLamaFiles);
        final List<Uri> fotoBaruFinal = new ArrayList<>(fotoBonUris);

        findViewById(R.id.btnEksporExcel).setEnabled(false);
        txtStatus.setText("Sedang membuat file Excel...");

        executor.execute(() -> {
            Uri outputUri = null;

            try {
                List<byte[]> fotoBytes = new ArrayList<>();
                int fotoGagal = 0;

                for (File f : fotoLamaFinal) {
                    try {
                        fotoBytes.add(bacaFile(f));
                    } catch (Exception e) {
                        fotoGagal++;
                        Log.w(TAG, "Foto lama tidak terbaca: " + f, e);
                    }
                }

                for (Uri fotoUri : fotoBaruFinal) {
                    try {
                        fotoBytes.add(bacaFoto(fotoUri));
                    } catch (Exception e) {
                        fotoGagal++;
                        Log.w(TAG, "Foto galeri tidak terbaca: " + fotoUri, e);
                    }
                }

                String projectFileName = proyekFinal
                        .replaceAll("[\\\\/:*?\"<>|]", "_")
                        .replaceAll("\\s+", "_")
                        .trim();

                String timestampSingkat = new SimpleDateFormat(
                        "yyMMdd", Locale.US).format(new Date());

                String fileName = "Reimbursement_" + projectFileName
                        + "_" + timestampSingkat + ".xlsx";

                Map<String, Map<String, Object>> allData = new LinkedHashMap<>();

                // FITUR 3: hanya isi sel input. Sel total & saldo
                // (H15-H18, H22) dibiarkan berupa rumus template.
                Map<String, Object> resume = new LinkedHashMap<>();
                resume.put("G10", proyekFinal);
                resume.put("G11", tanggalFinal);
                resume.put("G12", picFinal);
                resume.put("H20", kasFinal); // FITUR 2
                allData.put("Resume", resume);

                allData.put("Rincian", buatDataRincian(transaksiFinal));

                ContentValues values = new ContentValues();
                values.put(MediaStore.MediaColumns.DISPLAY_NAME, fileName);
                values.put(MediaStore.MediaColumns.MIME_TYPE, MIME_XLSX);
                values.put(MediaStore.MediaColumns.RELATIVE_PATH, FOLDER_PATH);
                values.put(MediaStore.MediaColumns.IS_PENDING, 1);

                Uri collection = MediaStore.Files.getContentUri(
                        MediaStore.VOLUME_EXTERNAL_PRIMARY);

                outputUri = getContentResolver().insert(collection, values);

                if (outputUri == null) {
                    throw new Exception("Gagal membuat file Excel.");
                }

                try (InputStream template = getAssets().open("TEMPLATE.xlsx");
                     OutputStream output =
                             getContentResolver().openOutputStream(outputUri)) {

                    if (output == null) {
                        throw new Exception("Gagal membuka file output.");
                    }

                    ExcelHelper.writeWorkbook(template, output, allData, fotoBytes);
                }

                // Pastikan hasil bisa dibaca sebelum dipublikasikan.
                try (InputStream verify =
                             getContentResolver().openInputStream(outputUri)) {

                    if (verify == null) {
                        throw new Exception("Gagal membaca hasil Excel.");
                    }

                    String hasil = ExcelHelper.readCell(verify, "Resume", "G10");

                    if (!proyekFinal.equals(hasil)) {
                        throw new Exception("Verifikasi Resume!G10 gagal: " + hasil);
                    }
                }

                ContentValues published = new ContentValues();
                published.put(MediaStore.MediaColumns.IS_PENDING, 0);
                getContentResolver().update(outputUri, published, null, null);

                final int gagal = fotoGagal;

                runOnUiThread(() -> {
                    if (isFinishing() || isDestroyed()) return;

                    String pesan = "Ekspor berhasil!\n" + FOLDER_PATH + fileName;

                    if (gagal > 0) {
                        pesan += "\nPerhatian: " + gagal
                                + " foto tidak bisa dibaca dan tidak ikut terlampir.";
                    }

                    txtStatus.setText(pesan);

                    Toast.makeText(this, "Excel berhasil disimpan",
                            Toast.LENGTH_LONG).show();

                    findViewById(R.id.btnEksporExcel).setEnabled(true);
                });

            } catch (Throwable t) {
                Log.e(TAG, "Gagal ekspor Excel", t);

                if (outputUri != null) {
                    try {
                        getContentResolver().delete(outputUri, null, null);
                    } catch (Throwable ignored) {
                        // Abaikan.
                    }
                }

                runOnUiThread(() -> {
                    if (isFinishing() || isDestroyed()) return;

                    findViewById(R.id.btnEksporExcel).setEnabled(true);
                    tampilkanError(
                            t instanceof OutOfMemoryError
                                    ? "Memori tidak cukup untuk membuat file"
                                    : "Ekspor gagal",
                            t);
                });
            }
        });
    }

    /**
     * Mengisi sheet Rincian.
     * A = nomor, B = tanggal, C = nominal, D = kategori, E = uraian.
     *
     * Material    : baris 9-37   (total di C38 tetap rumus template)
     * Jasa        : baris 43-49  (total di C50 tetap rumus template)
     * Operasional : baris 55-117 (total di C118 tetap rumus template)
     */
    private Map<String, Object> buatDataRincian(List<Transaksi> data)
            throws Exception {

        Map<String, Object> cells = new LinkedHashMap<>();

        int materialRow = 9;
        int jasaRow = 43;
        int operasionalRow = 55;
        int nomorMaterial = 1;
        int nomorJasa = 1;
        int nomorOperasional = 1;

        for (Transaksi t : data) {
            int row;

            switch (t.jenis) {
                case "Material":
                    if (materialRow > 37) {
                        throw new Exception(
                                "Transaksi Material maksimal 29 baris (baris 9-37).");
                    }
                    row = materialRow++;
                    cells.put("A" + row, nomorMaterial++);
                    break;

                case "Jasa":
                    if (jasaRow > 49) {
                        throw new Exception(
                                "Transaksi Jasa maksimal 7 baris (baris 43-49).");
                    }
                    row = jasaRow++;
                    cells.put("A" + row, nomorJasa++);
                    break;

                default:
                    if (operasionalRow > 117) {
                        throw new Exception(
                                "Transaksi Operasional maksimal 63 baris (baris 55-117).");
                    }
                    row = operasionalRow++;
                    cells.put("A" + row, nomorOperasional++);
                    break;
            }

            cells.put("B" + row, t.tanggal);
            cells.put("C" + row, t.nominal);
            cells.put("D" + row, t.kategori);
            cells.put("E" + row, t.uraian);
        }

        return cells;
    }

    // =====================================================================
    // FITUR 5: Draf otomatis (tidak hilang saat layar diputar / app ditutup)
    // =====================================================================

    private File folderFotoDraf() {
        File folder = new File(getFilesDir(), DRAFT_FOTO_DIR);
        if (!folder.exists()) folder.mkdirs();
        return folder;
    }

    private void simpanDraf() {
        try {
            JSONObject root = new JSONObject();
            root.put("proyek", editProject.getText().toString());
            root.put("pic", editPIC.getText().toString());
            root.put("tanggalReimburse", editReimburseDate.getText().toString());
            root.put("kas", editKas.getText().toString());
            root.put("jenis", jenisTerpilih);

            JSONArray daftar = new JSONArray();
            for (Transaksi t : transaksiList) {
                JSONObject o = new JSONObject();
                o.put("jenis", t.jenis);
                o.put("tanggal", t.tanggal);
                o.put("kategori", t.kategori);
                o.put("uraian", t.uraian);
                o.put("nominal", t.nominal);
                daftar.put(o);
            }
            root.put("transaksi", daftar);

            JSONArray uris = new JSONArray();
            for (Uri u : fotoBonUris) uris.put(u.toString());
            root.put("fotoUris", uris);

            JSONArray fotoLama = new JSONArray();
            for (File f : fotoLamaFiles) fotoLama.put(f.getName());
            root.put("fotoLama", fotoLama);

            // Tulis ke file sementara dulu, baru diganti, agar draf
            // tidak rusak jika aplikasi tertutup saat menyimpan.
            File tujuan = new File(getFilesDir(), DRAFT_FILE);
            File sementara = new File(getFilesDir(), DRAFT_FILE + ".tmp");
            tulisFile(sementara, root.toString().getBytes(StandardCharsets.UTF_8));

            if (!sementara.renameTo(tujuan)) {
                tulisFile(tujuan, root.toString().getBytes(StandardCharsets.UTF_8));
                sementara.delete();
            }
        } catch (Exception e) {
            Log.w(TAG, "Gagal menyimpan draf", e);
        }
    }

    private void muatDraf() {
        File file = new File(getFilesDir(), DRAFT_FILE);

        if (!file.exists()) {
            // Pertama kali dipakai: isi kas dengan nilai bawaan.
            editKas.setText(angkaUntukInput(KAS_DEFAULT));
            return;
        }

        try {
            JSONObject root = new JSONObject(
                    new String(bacaFile(file), StandardCharsets.UTF_8));

            editProject.setText(root.optString("proyek"));
            editPIC.setText(root.optString("pic"));
            editReimburseDate.setText(root.optString("tanggalReimburse"));
            editKas.setText(root.optString("kas"));

            jenisTerpilih = root.optString("jenis", "Material");
            perbaruiKategori(null);

            transaksiList.clear();
            JSONArray daftar = root.optJSONArray("transaksi");
            if (daftar != null) {
                for (int i = 0; i < daftar.length(); i++) {
                    JSONObject o = daftar.getJSONObject(i);
                    transaksiList.add(new Transaksi(
                            o.optString("jenis"),
                            o.optString("tanggal"),
                            o.optString("kategori"),
                            o.optString("uraian"),
                            o.optDouble("nominal", 0)));
                }
            }

            fotoBonUris.clear();
            JSONArray uris = root.optJSONArray("fotoUris");
            if (uris != null) {
                for (int i = 0; i < uris.length(); i++) {
                    fotoBonUris.add(Uri.parse(uris.getString(i)));
                }
            }

            fotoLamaFiles.clear();
            JSONArray fotoLama = root.optJSONArray("fotoLama");
            if (fotoLama != null) {
                File folder = folderFotoDraf();
                for (int i = 0; i < fotoLama.length(); i++) {
                    File f = new File(folder, fotoLama.getString(i));
                    if (f.exists()) fotoLamaFiles.add(f);
                }
            }
        } catch (Exception e) {
            Log.w(TAG, "Draf rusak, mulai dari kosong", e);
            editKas.setText(angkaUntukInput(KAS_DEFAULT));
        }
    }

    private void konfirmasiMulaiBaru() {
        new AlertDialog.Builder(this)
                .setTitle("Mulai reimbursement baru?")
                .setMessage("Semua data di layar (proyek, transaksi, dan lampiran foto) "
                        + "akan dikosongkan. File Excel yang sudah diekspor tidak terpengaruh.")
                .setPositiveButton("Kosongkan", (d, w) -> mulaiBaru())
                .setNegativeButton("Batal", null)
                .show();
    }

    private void mulaiBaru() {
        keluarModeEdit(true);
        kosongkanFormTransaksi();

        editProject.setText("");
        editPIC.setText("");
        editReimburseDate.setText("");
        editKas.setText(angkaUntukInput(KAS_DEFAULT));

        transaksiList.clear();
        hapusSemuaFoto();
        pilihJenis("Material");

        perbaruiRekap();
        simpanDraf();

        txtStatus.setText("Siap digunakan.");
    }

    // =====================================================================
    // Utilitas
    // =====================================================================

    /**
     * Menampilkan pesan error beserta detailnya. Detail bisa disalin
     * supaya mudah dilaporkan tanpa perlu membuka Logcat.
     */
    private void tampilkanError(String judul, Throwable t) {
        String ringkas = t.getClass().getSimpleName()
                + (t.getMessage() != null ? ": " + t.getMessage() : "");

        txtStatus.setText(judul + "\n" + ringkas);

        StringBuilder detail = new StringBuilder(t.toString());
        StackTraceElement[] jejak = t.getStackTrace();
        for (int i = 0; i < Math.min(jejak.length, 12); i++) {
            detail.append("\n  at ").append(jejak[i]);
        }

        Throwable sebab = t.getCause();
        if (sebab != null && sebab != t) {
            detail.append("\nCaused by: ").append(sebab);
            StackTraceElement[] jejakSebab = sebab.getStackTrace();
            for (int i = 0; i < Math.min(jejakSebab.length, 6); i++) {
                detail.append("\n  at ").append(jejakSebab[i]);
            }
        }

        final String teksDetail = detail.toString();

        new AlertDialog.Builder(this)
                .setTitle(judul)
                .setMessage(teksDetail)
                .setPositiveButton("Salin Detail", (d, w) -> {
                    ClipboardManager clipboard =
                            (ClipboardManager) getSystemService(CLIPBOARD_SERVICE);
                    if (clipboard != null) {
                        clipboard.setPrimaryClip(
                                ClipData.newPlainText("Error Rembuse", teksDetail));
                        Toast.makeText(this, "Detail error disalin",
                                Toast.LENGTH_SHORT).show();
                    }
                })
                .setNegativeButton("Tutup", null)
                .show();
    }

    private static String teksAman(String s) {
        return s == null ? "" : s;
    }

    private static void salinKeFile(InputStream input, File tujuan) throws Exception {
        try (OutputStream output = new FileOutputStream(tujuan)) {
            byte[] buffer = new byte[8192];
            int jumlah;
            while ((jumlah = input.read(buffer)) != -1) {
                output.write(buffer, 0, jumlah);
            }
        }
    }

    private double ambilKas() {
        double kas = parseAngka(editKas.getText().toString());
        return kas > 0 ? kas : 0;
    }

    /**
     * Mengubah teks angka menjadi double.
     * Contoh: "150000", "150000.5", "150000,5", "Rp150.000", " 3,000,000 ".
     * Mengembalikan 0 jika tidak bisa dibaca.
     */
    static double parseAngka(String teks) {
        if (teks == null) return 0;

        String s = teks.trim().replaceAll("[^0-9.,-]", "");
        if (s.isEmpty() || s.equals("-")) return 0;

        int titik = s.lastIndexOf('.');
        int koma = s.lastIndexOf(',');

        if (titik >= 0 && koma >= 0) {
            // Pemisah yang muncul terakhir adalah desimal.
            if (koma > titik) {
                s = s.replace(".", "").replace(',', '.');
            } else {
                s = s.replace(",", "");
            }
        } else if (titik >= 0 || koma >= 0) {
            char sep = titik >= 0 ? '.' : ',';
            String[] bagian = s.split(sep == '.' ? "\\." : ",", -1);

            // Satu pemisah dengan tepat 3 digit di belakang (150.000)
            // atau lebih dari satu pemisah = pemisah ribuan.
            boolean ribuan = bagian.length > 2
                    || (bagian.length == 2 && bagian[1].length() == 3);

            s = ribuan
                    ? s.replace(String.valueOf(sep), "")
                    : s.replace(sep, '.');
        }

        try {
            double hasil = Double.parseDouble(s);
            return Double.isFinite(hasil) ? hasil : 0;
        } catch (NumberFormatException e) {
            return 0;
        }
    }

    /** Angka tanpa ".0" untuk ditampilkan di kolom input. */
    private static String angkaUntukInput(double angka) {
        if (angka == Math.rint(angka) && Math.abs(angka) < 1e15) {
            return String.valueOf((long) angka);
        }
        return BigDecimal.valueOf(angka).stripTrailingZeros().toPlainString();
    }

    private String formatRupiah(double angka) {
        NumberFormat rupiah = NumberFormat.getCurrencyInstance(localeIndonesia);
        rupiah.setMaximumFractionDigits(0);
        return rupiah.format(angka);
    }

    private double totalDari(List<Transaksi> data, String jenis) {
        double total = 0;
        for (Transaksi t : data) {
            if (t.jenis.equals(jenis)) total += t.nominal;
        }
        return total;
    }

    private int dp(int nilai) {
        return Math.round(nilai * getResources().getDisplayMetrics().density);
    }

    private byte[] bacaFoto(Uri uri) throws Exception {
        try (InputStream input = getContentResolver().openInputStream(uri)) {
            if (input == null) {
                throw new Exception("Foto tidak dapat dibaca: " + uri);
            }
            return bacaSemua(input);
        }
    }

    private static byte[] bacaFile(File file) throws Exception {
        try (InputStream input = new FileInputStream(file)) {
            return bacaSemua(input);
        }
    }

    private static byte[] bacaSemua(InputStream input) throws Exception {
        try (ByteArrayOutputStream output = new ByteArrayOutputStream()) {
            byte[] buffer = new byte[8192];
            int jumlah;
            while ((jumlah = input.read(buffer)) != -1) {
                output.write(buffer, 0, jumlah);
            }
            return output.toByteArray();
        }
    }

    private static void tulisFile(File file, byte[] data) throws Exception {
        try (OutputStream output = new FileOutputStream(file)) {
            output.write(data);
        }
    }

    private static void hapusIsiFolder(File folder) {
        File[] isi = folder.listFiles();
        if (isi == null) return;
        for (File f : isi) {
            if (f.isFile()) f.delete();
        }
    }

    // =====================================================================
    // Model data
    // =====================================================================

    private static class Transaksi {
        final String jenis;
        final String tanggal;
        final String kategori;
        final String uraian;
        final double nominal;

        Transaksi(String jenis, String tanggal, String kategori,
                  String uraian, double nominal) {
            this.jenis = jenis;
            this.tanggal = tanggal;
            this.kategori = kategori;
            this.uraian = uraian;
            this.nominal = nominal;
        }
    }
}