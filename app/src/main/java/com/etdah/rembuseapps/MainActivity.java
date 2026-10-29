package com.etdah.rembuseapps;

import android.app.DatePickerDialog;
import android.content.ClipData;
import android.content.ClipboardManager;
import android.content.ContentValues;
import android.content.Intent;
import android.graphics.Color;
import android.graphics.Typeface;
import android.net.Uri;
import android.os.Bundle;
import android.provider.MediaStore;
import android.text.Editable;
import android.text.TextWatcher;
import android.util.Log;
import android.util.TypedValue;
import android.view.View;
import android.widget.AdapterView;
import android.widget.ArrayAdapter;
import android.widget.Button;
import android.widget.EditText;
import android.widget.LinearLayout;
import android.widget.ScrollView;
import android.widget.Spinner;
import android.widget.TextView;
import android.widget.Toast;

import androidx.activity.result.ActivityResultLauncher;
import androidx.activity.result.contract.ActivityResultContracts;
import androidx.appcompat.app.AlertDialog;
import androidx.appcompat.app.AppCompatActivity;
import androidx.core.content.FileProvider;

import com.google.android.material.badge.BadgeDrawable;
import com.google.android.material.bottomnavigation.BottomNavigationView;

import org.json.JSONArray;
import org.json.JSONObject;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.text.NumberFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Collections;
import java.util.Date;
import java.util.HashSet;
import java.util.IdentityHashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class MainActivity extends AppCompatActivity {

    private static final String TAG = "REMBUSE";

    private static final String FOLDER_PATH = "Documents/Rembuse/";

    private static final String MIME_XLSX =
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

    /** File draf agar data tidak hilang saat layar diputar / app ditutup. */
    private static final String DRAFT_FILE = "draft.json";

    /** Folder foto bon yang sudah dikompres (memori internal aplikasi). */
    private static final String FOLDER_FOTO = "foto";

    /** Folder foto versi sebelumnya (untuk migrasi). */
    private static final String FOLDER_FOTO_LAMA = "draft_foto";

    private static final String KEY_EDIT_INDEX = "editIndex";
    private static final String KEY_TUJUAN_FOTO = "tujuanFoto";
    private static final String KEY_HALAMAN = "halaman";

    private static final String TUJUAN_FORM = "form";
    private static final String TUJUAN_UMUM = "umum";

    /** Keterangan foto di sheet Lampiran, misalnya "Bon Material no. 2 | ...". */
    private static final Pattern POLA_KETERANGAN_BON =
            Pattern.compile("^Bon (Material|Jasa|Operasional) no\\.\\s*(\\d+)");

    private static final String[] PILIHAN_FILTER = {
            "Semua jenis", "Material", "Jasa", "Operasional", "Ada bon", "Tanpa bon"
    };

    private static final String[] PILIHAN_URUT = {
            "Urutan input", "Tanggal terbaru", "Tanggal terlama",
            "Nominal terbesar", "Nominal terkecil"
    };

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
    private Spinner spinnerKategori, spinnerFilter, spinnerUrut;
    private TextView txtJenisTerpilih, txtFotoForm;
    private TextView txtTotalSemua, txtSaldo;
    private TextView txtDaftarTransaksi, txtStatus, txtJumlahFoto;
    private LinearLayout listTransaksi;
    private RingkasanView ringkasanView;
    private TextView txtNamaPerusahaan;
    private ScrollView scrollUtama;
    private BottomNavigationView navigasiBawah;
    private View halamanInput, halamanRekap, halamanFile;
    private Button btnMaterial, btnJasa, btnOperasional;
    private Button btnTambahTransaksi, btnBatalEdit, btnBagikan;

    // ---------- Data ----------
    private String jenisTerpilih = "Material";
    private final List<Transaksi> transaksiList = new ArrayList<>();

    /** Foto bon untuk transaksi yang sedang diisi / diedit (nama file). */
    private final List<String> fotoForm = new ArrayList<>();

    /** Lampiran lain yang tidak terkait transaksi tertentu (nama file). */
    private final List<String> fotoUmum = new ArrayList<>();

    /** Posisi transaksi yang sedang diedit, -1 jika sedang menambah baru. */
    private int editIndex = -1;

    /** Ke mana foto dari kamera/galeri akan dimasukkan. */
    private String tujuanFoto = TUJUAN_FORM;

    /** File Excel terakhir yang diekspor (untuk tombol Bagikan). */
    private String terakhirEksporUri = "";
    private String terakhirEksporNama = "";

    private RiwayatEkspor riwayat;
    private Pengaturan pengaturan;

    /** Tab yang sedang tampil (id item menu navigasi). */
    private int halamanAktif;

    private ActivityResultLauncher<String[]> reopenExcelPicker;
    private ActivityResultLauncher<String[]> galeriPicker;
    private ActivityResultLauncher<Uri> kameraLauncher;

    // =====================================================================
    // Siklus hidup activity
    // =====================================================================

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        riwayat = new RiwayatEkspor(getFilesDir());
        pengaturan = new Pengaturan(this);
        halamanAktif = R.id.nav_input;

        if (savedInstanceState != null) {
            tujuanFoto = savedInstanceState.getString(KEY_TUJUAN_FOTO, TUJUAN_FORM);
            halamanAktif = savedInstanceState.getInt(KEY_HALAMAN, R.id.nav_input);
        }

        hubungkanKomponen();
        aturNavigasi();
        terapkanNamaPerusahaan();
        daftarkanLauncher();
        aturTombol();
        aturPemilihTanggal();
        aturKategori();
        aturFilterDanUrutan();

        // Pulihkan data terakhir (setelah layar diputar / app ditutup).
        muatDraf();

        // Format ribuan otomatis (150000 -> 150.000).
        editNominal.addTextChangedListener(new FormatRibuan(editNominal, null));
        editKas.addTextChangedListener(new FormatRibuan(editKas, this::perbaruiRekap));

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

        bersihkanFotoTakTerpakai();

        perbaruiTombolJenis();
        perbaruiFotoForm();
        perbaruiJumlahFoto();
        perbaruiTombolBagikan();
        perbaruiRekap();
    }

    @Override
    protected void onSaveInstanceState(Bundle outState) {
        super.onSaveInstanceState(outState);
        outState.putInt(KEY_EDIT_INDEX, editIndex);
        outState.putString(KEY_TUJUAN_FOTO, tujuanFoto);
        outState.putInt(KEY_HALAMAN, halamanAktif);
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
        spinnerFilter = findViewById(R.id.spinnerFilter);
        spinnerUrut = findViewById(R.id.spinnerUrut);

        txtJenisTerpilih = findViewById(R.id.txtJenisTerpilih);
        txtFotoForm = findViewById(R.id.txtFotoForm);
        txtTotalSemua = findViewById(R.id.txtTotalSemua);
        txtSaldo = findViewById(R.id.txtSaldo);
        txtDaftarTransaksi = findViewById(R.id.txtDaftarTransaksi);
        txtStatus = findViewById(R.id.txtStatus);
        txtJumlahFoto = findViewById(R.id.txtJumlahFoto);
        listTransaksi = findViewById(R.id.listTransaksi);
        ringkasanView = findViewById(R.id.ringkasanView);
        txtNamaPerusahaan = findViewById(R.id.txtNamaPerusahaan);
        scrollUtama = findViewById(R.id.scrollUtama);
        navigasiBawah = findViewById(R.id.navigasiBawah);
        halamanInput = findViewById(R.id.halamanInput);
        halamanRekap = findViewById(R.id.halamanRekap);
        halamanFile = findViewById(R.id.halamanFile);

        btnMaterial = findViewById(R.id.btnMaterial);
        btnJasa = findViewById(R.id.btnJasa);
        btnOperasional = findViewById(R.id.btnOperasional);
        btnTambahTransaksi = findViewById(R.id.btnTambahTransaksi);
        btnBatalEdit = findViewById(R.id.btnBatalEdit);
        btnBagikan = findViewById(R.id.btnBagikan);
    }

    private void daftarkanLauncher() {
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

        galeriPicker = registerForActivityResult(
                new ActivityResultContracts.OpenMultipleDocuments(),
                uris -> {
                    if (uris == null || uris.isEmpty()) return;
                    prosesFotoGaleri(uris, tujuanFoto);
                }
        );

        kameraLauncher = registerForActivityResult(
                new ActivityResultContracts.TakePicture(),
                berhasil -> {
                    if (berhasil == null || !berhasil) return;
                    prosesFotoKamera(tujuanFoto);
                }
        );
    }

    private void aturTombol() {
        btnTambahTransaksi.setOnClickListener(v -> simpanTransaksi());
        btnBatalEdit.setOnClickListener(v -> keluarModeEdit(true));

        // Foto bon untuk transaksi (opsional).
        findViewById(R.id.btnFotoKamera).setOnClickListener(v -> bukaKamera(TUJUAN_FORM));
        findViewById(R.id.btnFotoGaleri).setOnClickListener(v -> bukaGaleri(TUJUAN_FORM));
        findViewById(R.id.btnFotoHapusForm).setOnClickListener(v -> lepasFotoForm());

        // Lampiran lain.
        findViewById(R.id.btnKameraLain).setOnClickListener(v -> bukaKamera(TUJUAN_UMUM));
        findViewById(R.id.btnPilihFoto).setOnClickListener(v -> bukaGaleri(TUJUAN_UMUM));
        findViewById(R.id.btnHapusFoto).setOnClickListener(v -> konfirmasiHapusFotoUmum());

        // File.
        findViewById(R.id.btnReopenExcel).setOnClickListener(v ->
                konfirmasiGantiData(() ->
                        reopenExcelPicker.launch(new String[]{MIME_XLSX})));

        findViewById(R.id.btnRiwayat).setOnClickListener(v -> tampilkanRiwayat());
        findViewById(R.id.btnEksporExcel).setOnClickListener(v -> eksporExcel());

        btnBagikan.setOnClickListener(v -> {
            if (!terakhirEksporUri.isEmpty()) {
                bagikanFile(Uri.parse(terakhirEksporUri), terakhirEksporNama);
            }
        });

        findViewById(R.id.btnMulaiBaru).setOnClickListener(v -> konfirmasiMulaiBaru());
        findViewById(R.id.btnPengaturan).setOnClickListener(v -> tampilkanPengaturan());
    }

    // =====================================================================
    // FITUR 10: Tab Input / Rekap / File
    // =====================================================================

    private void aturNavigasi() {
        navigasiBawah.setOnItemSelectedListener(item -> {
            tampilkanHalaman(item.getItemId());
            return true;
        });

        navigasiBawah.setSelectedItemId(halamanAktif);
        tampilkanHalaman(halamanAktif);
    }

    private void tampilkanHalaman(int idMenu) {
        // Catatan: id resource tidak boleh dipakai di switch (AGP 8+),
        // jadi memakai if-else.
        if (idMenu != R.id.nav_rekap && idMenu != R.id.nav_file) {
            idMenu = R.id.nav_input;
        }

        boolean berubah = idMenu != halamanAktif;
        halamanAktif = idMenu;

        halamanInput.setVisibility(idMenu == R.id.nav_input ? View.VISIBLE : View.GONE);
        halamanRekap.setVisibility(idMenu == R.id.nav_rekap ? View.VISIBLE : View.GONE);
        halamanFile.setVisibility(idMenu == R.id.nav_file ? View.VISIBLE : View.GONE);

        if (berubah) scrollUtama.scrollTo(0, 0);
    }

    /** Pindah tab dari kode (misalnya saat mengedit dari halaman Rekap). */
    private void pindahHalaman(int idMenu) {
        if (navigasiBawah.getSelectedItemId() != idMenu) {
            navigasiBawah.setSelectedItemId(idMenu);
        } else {
            tampilkanHalaman(idMenu);
        }
    }

    /** Angka jumlah transaksi di ikon tab Rekap. */
    private void perbaruiBadgeRekap() {
        BadgeDrawable badge = navigasiBawah.getOrCreateBadge(R.id.nav_rekap);
        badge.setVisible(!transaksiList.isEmpty());
        badge.setNumber(transaksiList.size());
        badge.setBackgroundColor(Color.parseColor("#6741A5"));
    }

    // =====================================================================
    // FITUR 8: Pengaturan
    // =====================================================================

    private void terapkanNamaPerusahaan() {
        String nama = pengaturan.namaPerusahaan();
        txtNamaPerusahaan.setText(
                (nama.isEmpty() ? "REMBUSE" : nama.toUpperCase(localeIndonesia))
                        + "  •  REIMBURSEMENT");
    }

    private void tampilkanPengaturan() {
        View tampilan = getLayoutInflater().inflate(R.layout.dialog_pengaturan, null);

        EditText editPerusahaan = tampilan.findViewById(R.id.editSetPerusahaan);
        EditText editPicDefault = tampilan.findViewById(R.id.editSetPIC);
        EditText editKasDefault = tampilan.findViewById(R.id.editSetKas);
        EditText editKatMaterial = tampilan.findViewById(R.id.editSetKatMaterial);
        EditText editKatJasa = tampilan.findViewById(R.id.editSetKatJasa);
        EditText editKatOperasional = tampilan.findViewById(R.id.editSetKatOperasional);

        editPerusahaan.setText(pengaturan.namaPerusahaan());
        editPicDefault.setText(pengaturan.picDefault());
        editKasDefault.setText(angkaUntukInput(pengaturan.kasDefault()));
        editKatMaterial.setText(pengaturan.kategoriTeks("Material"));
        editKatJasa.setText(pengaturan.kategoriTeks("Jasa"));
        editKatOperasional.setText(pengaturan.kategoriTeks("Operasional"));

        editKasDefault.addTextChangedListener(new FormatRibuan(editKasDefault, null));

        new AlertDialog.Builder(this)
                .setTitle("Pengaturan")
                .setView(tampilan)
                .setPositiveButton("Simpan", (d, w) -> {
                    pengaturan.simpan(
                            editPerusahaan.getText().toString(),
                            editPicDefault.getText().toString(),
                            parseAngka(editKasDefault.getText().toString()),
                            editKatMaterial.getText().toString(),
                            editKatJasa.getText().toString(),
                            editKatOperasional.getText().toString());
                    setelahPengaturanBerubah();
                    Toast.makeText(this, "Pengaturan disimpan", Toast.LENGTH_SHORT).show();
                })
                .setNeutralButton("Kembalikan Bawaan", (d, w) ->
                        new AlertDialog.Builder(this)
                                .setTitle("Kembalikan pengaturan bawaan?")
                                .setMessage("Nama perusahaan, PIC default, kas default, "
                                        + "dan kategori kembali seperti semula. "
                                        + "Transaksi yang sudah ada tidak berubah.")
                                .setPositiveButton("Kembalikan", (d2, w2) -> {
                                    pengaturan.kembalikanBawaan();
                                    setelahPengaturanBerubah();
                                })
                                .setNegativeButton("Batal", null)
                                .show())
                .setNegativeButton("Batal", null)
                .show();
    }

    private void setelahPengaturanBerubah() {
        terapkanNamaPerusahaan();

        // Pertahankan kategori yang sedang dipilih jika masih ada.
        Object terpilih = spinnerKategori.getSelectedItem();
        perbaruiKategori(editIndex >= 0 && terpilih != null ? terpilih.toString() : null);

        if (editPIC.getText().toString().trim().isEmpty()) {
            editPIC.setText(pengaturan.picDefault());
        }
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
        Date tanggal = parseTanggal(target.getText().toString());
        if (tanggal != null) calendar.setTime(tanggal);

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

    private Date parseTanggal(String teks) {
        if (teks == null || teks.trim().isEmpty()) return null;
        try {
            return dateFormat.parse(teks.trim());
        } catch (Exception e) {
            return null;
        }
    }

    // =====================================================================
    // Jenis, kategori & kapasitas (FITUR 9)
    // =====================================================================

    /** Jumlah baris yang disediakan template Excel untuk tiap jenis. */
    private static int kapasitas(String jenis) {
        switch (jenis) {
            case "Material":
                return 29;  // baris 9-37
            case "Jasa":
                return 7;   // baris 43-49
            default:
                return 63;  // baris 55-117
        }
    }

    private int jumlahJenis(String jenis, int kecualiIndex) {
        int jumlah = 0;
        for (int i = 0; i < transaksiList.size(); i++) {
            if (i != kecualiIndex && transaksiList.get(i).jenis.equals(jenis)) jumlah++;
        }
        return jumlah;
    }

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
        return pengaturan.kategori(jenis);
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
        btnMaterial.setAlpha("Material".equals(jenisTerpilih) ? 1.0f : 0.5f);
        btnJasa.setAlpha("Jasa".equals(jenisTerpilih) ? 1.0f : 0.5f);
        btnOperasional.setAlpha("Operasional".equals(jenisTerpilih) ? 1.0f : 0.5f);
        perbaruiInfoJenis();
    }

    /** Menampilkan jenis terpilih beserta sisa kapasitas barisnya. */
    private void perbaruiInfoJenis() {
        int kap = kapasitas(jenisTerpilih);
        int terpakai = jumlahJenis(jenisTerpilih, -1);
        int sisa = kap - terpakai;

        String teks = "Jenis: " + jenisTerpilih + "   •   " + terpakai + "/" + kap + " baris";
        int warna;

        if (sisa <= 0) {
            teks += "  (penuh)";
            warna = Color.parseColor("#B3261E");
        } else if (sisa <= 3) {
            teks += "  (sisa " + sisa + ")";
            warna = Color.parseColor("#B26A00");
        } else {
            warna = Color.parseColor("#6741A5");
        }

        txtJenisTerpilih.setText(teks);
        txtJenisTerpilih.setTextColor(warna);
    }

    // =====================================================================
    // FITUR 1, 3, 4: Foto bon (kompres, kamera, per transaksi)
    // =====================================================================

    private File folderFoto() {
        File folder = new File(getFilesDir(), FOLDER_FOTO);
        if (!folder.exists()) folder.mkdirs();
        return folder;
    }

    private File fileFoto(String nama) {
        return new File(folderFoto(), nama);
    }

    private File fileKamera() {
        File folder = new File(getCacheDir(), "kamera");
        if (!folder.exists()) folder.mkdirs();
        return new File(folder, "foto_kamera.jpg");
    }

    private void bukaKamera(String tujuan) {
        tujuanFoto = tujuan;

        try {
            File file = fileKamera();
            if (file.exists()) file.delete();

            Uri uri = FileProvider.getUriForFile(
                    this, getPackageName() + ".fileprovider", file);

            kameraLauncher.launch(uri);
        } catch (Exception e) {
            Log.e(TAG, "Kamera tidak bisa dibuka", e);
            Toast.makeText(this, "Aplikasi kamera tidak tersedia",
                    Toast.LENGTH_LONG).show();
        }
    }

    private void bukaGaleri(String tujuan) {
        tujuanFoto = tujuan;

        try {
            galeriPicker.launch(new String[]{"image/*"});
        } catch (Exception e) {
            Toast.makeText(this, "Galeri tidak bisa dibuka",
                    Toast.LENGTH_LONG).show();
        }
    }

    private void prosesFotoKamera(String tujuan) {
        txtStatus.setText("Mengompres foto...");

        executor.execute(() -> {
            try {
                File sumber = fileKamera();
                String nama = FotoHelper.kompresFile(sumber, folderFoto());
                sumber.delete();

                List<String> hasil = new ArrayList<>();
                hasil.add(nama);

                runOnUiThread(() -> tambahFoto(tujuan, hasil, 0));
            } catch (Throwable t) {
                Log.e(TAG, "Foto kamera gagal diproses", t);
                runOnUiThread(() -> {
                    if (isFinishing() || isDestroyed()) return;
                    tampilkanError("Foto gagal diproses", t);
                });
            }
        });
    }

    private void prosesFotoGaleri(List<Uri> uris, String tujuan) {
        txtStatus.setText("Mengompres " + uris.size() + " foto...");

        executor.execute(() -> {
            List<String> hasil = new ArrayList<>();
            int gagal = 0;

            for (Uri uri : uris) {
                try (InputStream input = getContentResolver().openInputStream(uri)) {
                    if (input == null) throw new Exception("Foto tidak dapat dibaca.");
                    hasil.add(FotoHelper.simpanTerkompres(
                            input, getCacheDir(), folderFoto()));
                } catch (Throwable t) {
                    gagal++;
                    Log.w(TAG, "Foto galeri gagal: " + uri, t);
                }
            }

            final int jumlahGagal = gagal;
            runOnUiThread(() -> tambahFoto(tujuan, hasil, jumlahGagal));
        });
    }

    private void tambahFoto(String tujuan, List<String> namaFile, int gagal) {
        if (isFinishing() || isDestroyed()) return;

        if (TUJUAN_UMUM.equals(tujuan)) {
            fotoUmum.addAll(namaFile);
        } else {
            fotoForm.addAll(namaFile);
        }

        perbaruiFotoForm();
        perbaruiJumlahFoto();
        simpanDraf();

        String pesan = namaFile.size() + " foto ditambahkan"
                + (TUJUAN_UMUM.equals(tujuan) ? " ke lampiran lain." : " ke transaksi ini.");
        if (gagal > 0) pesan += " " + gagal + " foto gagal dibaca.";

        txtStatus.setText(pesan);
        Toast.makeText(this, pesan, Toast.LENGTH_SHORT).show();
    }

    private void perbaruiFotoForm() {
        if (fotoForm.isEmpty()) {
            txtFotoForm.setText("Foto bon (opsional): belum ada\n"
                    + "Transaksi tanpa bon tetap bisa disimpan.");
        } else {
            txtFotoForm.setText("Foto bon: " + fotoForm.size() + " foto terlampir");
        }

        findViewById(R.id.btnFotoHapusForm).setEnabled(!fotoForm.isEmpty());
    }

    private void lepasFotoForm() {
        if (fotoForm.isEmpty()) return;

        fotoForm.clear();
        perbaruiFotoForm();
        simpanDraf();

        Toast.makeText(this, "Foto dilepas dari transaksi ini",
                Toast.LENGTH_SHORT).show();
    }

    private void perbaruiJumlahFoto() {
        txtJumlahFoto.setText("Lampiran lain: " + fotoUmum.size() + " foto");
    }

    private void konfirmasiHapusFotoUmum() {
        if (fotoUmum.isEmpty()) {
            Toast.makeText(this, "Belum ada lampiran lain", Toast.LENGTH_SHORT).show();
            return;
        }

        new AlertDialog.Builder(this)
                .setTitle("Hapus lampiran lain?")
                .setMessage(fotoUmum.size() + " foto akan dilepas dari reimbursement ini. "
                        + "Foto bon pada transaksi dan foto asli di galeri tidak terpengaruh.")
                .setPositiveButton("Hapus", (d, w) -> {
                    fotoUmum.clear();
                    perbaruiJumlahFoto();
                    simpanDraf();
                })
                .setNegativeButton("Batal", null)
                .show();
    }

    /**
     * Menghapus file foto yang tidak lagi dipakai transaksi mana pun.
     * Dijalankan di antrean executor agar tidak bentrok dengan proses foto.
     */
    private void bersihkanFotoTakTerpakai() {
        final Set<String> dipakai = new HashSet<>(fotoForm);
        dipakai.addAll(fotoUmum);
        for (Transaksi t : transaksiList) dipakai.addAll(t.foto);

        final File folder = folderFoto();

        // Hanya file yang sudah ada SEBELUM saat ini yang boleh dihapus.
        // File yang dibuat sesudahnya (misalnya foto yang sedang diproses)
        // tetap aman walaupun belum tercatat di daftar.
        final long batasWaktu = System.currentTimeMillis() - 2000;

        executor.execute(() -> {
            File[] isi = folder.listFiles();
            if (isi == null) return;

            for (File f : isi) {
                if (f.isFile()
                        && !dipakai.contains(f.getName())
                        && f.lastModified() < batasWaktu) {
                    f.delete();
                }
            }
        });
    }

    // =====================================================================
    // Tambah, edit, dan hapus transaksi
    // =====================================================================

    /** Dipanggil tombol Tambah / Simpan Perubahan. */
    private void simpanTransaksi() {
        String tanggal = editPurchaseDate.getText().toString().trim();
        String uraian = editUraian.getText().toString().trim();
        double nominal = parseAngka(editNominal.getText().toString());

        if (tanggal.isEmpty()) {
            editPurchaseDate.setError("Tanggal pembelian wajib diisi");
            return;
        }

        if (uraian.isEmpty()) {
            editUraian.setError("Uraian wajib diisi");
            return;
        }

        if (!(nominal > 0)) {
            editNominal.setError("Masukkan nominal lebih besar dari nol");
            return;
        }

        boolean modeEdit = editIndex >= 0 && editIndex < transaksiList.size();

        // FITUR 9: cek kapasitas baris template.
        int kap = kapasitas(jenisTerpilih);
        int terpakai = jumlahJenis(jenisTerpilih, modeEdit ? editIndex : -1);

        if (terpakai >= kap) {
            new AlertDialog.Builder(this)
                    .setTitle("Baris " + jenisTerpilih + " sudah penuh")
                    .setMessage("Template Excel hanya menyediakan " + kap
                            + " baris untuk " + jenisTerpilih + ". "
                            + "Hapus atau gabungkan transaksi " + jenisTerpilih
                            + " yang lain, atau buat reimbursement terpisah.")
                    .setPositiveButton("Mengerti", null)
                    .show();
            return;
        }

        Object item = spinnerKategori.getSelectedItem();
        String kategori = item == null ? "" : item.toString();

        Transaksi transaksi = new Transaksi(
                jenisTerpilih, tanggal, kategori, uraian, nominal, fotoForm);

        if (modeEdit) {
            transaksiList.set(editIndex, transaksi);
        } else {
            transaksiList.add(transaksi);
        }

        keluarModeEdit(false);
        kosongkanFormTransaksi();
        perbaruiRekap();
        simpanDraf();

        String pesan = modeEdit ? "Transaksi diperbarui" : "Transaksi berhasil ditambahkan";
        int sisa = kap - jumlahJenis(jenisTerpilih, -1);
        if (sisa > 0 && sisa <= 3) {
            pesan += ". Sisa " + sisa + " baris " + jenisTerpilih;
        }

        Toast.makeText(this, pesan, Toast.LENGTH_SHORT).show();
    }

    private void kosongkanFormTransaksi() {
        editPurchaseDate.setText("");
        editUraian.setText("");
        editNominal.setText("");
        editPurchaseDate.setError(null);
        editUraian.setError(null);
        editNominal.setError(null);

        fotoForm.clear();
        perbaruiFotoForm();
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

        pindahHalaman(R.id.nav_input);

        jenisTerpilih = t.jenis;
        perbaruiKategori(t.kategori);
        perbaruiTombolJenis();

        editPurchaseDate.setText(t.tanggal);
        editUraian.setText(t.uraian);
        editNominal.setText(angkaUntukInput(t.nominal));

        fotoForm.clear();
        fotoForm.addAll(t.foto);
        perbaruiFotoForm();

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
                        + t.uraian + "\n" + formatRupiah(t.nominal)
                        + (t.foto.isEmpty() ? "" : "\n" + t.foto.size() + " foto bon ikut terhapus"))
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
    // Rekap, ringkasan visual (FITUR 11), filter & urutan (FITUR 10)
    // =====================================================================

    private void aturFilterDanUrutan() {
        ArrayAdapter<String> filter = new ArrayAdapter<>(
                this, android.R.layout.simple_spinner_item, PILIHAN_FILTER);
        filter.setDropDownViewResource(android.R.layout.simple_spinner_dropdown_item);
        spinnerFilter.setAdapter(filter);

        ArrayAdapter<String> urut = new ArrayAdapter<>(
                this, android.R.layout.simple_spinner_item, PILIHAN_URUT);
        urut.setDropDownViewResource(android.R.layout.simple_spinner_dropdown_item);
        spinnerUrut.setAdapter(urut);

        AdapterView.OnItemSelectedListener listener = new AdapterView.OnItemSelectedListener() {
            @Override
            public void onItemSelected(AdapterView<?> parent, View view, int position, long id) {
                tampilkanDaftarTransaksi();
            }

            @Override
            public void onNothingSelected(AdapterView<?> parent) {
            }
        };

        spinnerFilter.setOnItemSelectedListener(listener);
        spinnerUrut.setOnItemSelectedListener(listener);
    }

    private void perbaruiRekap() {
        double totalMaterial = totalDari(transaksiList, "Material");
        double totalJasa = totalDari(transaksiList, "Jasa");
        double totalOperasional = totalDari(transaksiList, "Operasional");
        double total = totalMaterial + totalJasa + totalOperasional;

        double kas = ambilKas();
        double saldo = kas - total;

        ringkasanView.setData(totalMaterial, totalJasa, totalOperasional, kas);
        txtTotalSemua.setText("TOTAL: " + formatRupiah(total));

        txtSaldo.setText("Kas: " + formatRupiah(kas)
                + "\nSaldo: " + formatRupiah(saldo)
                + (saldo < 0 ? "  (kurang)" : ""));
        txtSaldo.setTextColor(saldo < 0
                ? Color.parseColor("#B3261E")
                : Color.parseColor("#6741A5"));

        perbaruiInfoJenis();
        perbaruiBadgeRekap();
        tampilkanDaftarTransaksi();
    }

    private boolean lolosFilter(Transaksi t, int filter) {
        switch (filter) {
            case 1:
                return t.jenis.equals("Material");
            case 2:
                return t.jenis.equals("Jasa");
            case 3:
                return t.jenis.equals("Operasional");
            case 4:
                return !t.foto.isEmpty();
            case 5:
                return t.foto.isEmpty();
            default:
                return true;
        }
    }

    /** Urutan tampilan (indeks asli ke transaksiList). */
    private List<Integer> indeksTampil() {
        int filter = spinnerFilter.getSelectedItemPosition();
        int urut = spinnerUrut.getSelectedItemPosition();

        List<Integer> indeks = new ArrayList<>();
        for (int i = 0; i < transaksiList.size(); i++) {
            if (lolosFilter(transaksiList.get(i), filter)) indeks.add(i);
        }

        switch (urut) {
            case 1: // Tanggal terbaru
            case 2: // Tanggal terlama
                final boolean terbaru = urut == 1;
                Collections.sort(indeks, (a, b) -> {
                    Date da = parseTanggal(transaksiList.get(a).tanggal);
                    Date db = parseTanggal(transaksiList.get(b).tanggal);
                    // Tanggal yang tidak terbaca selalu di bawah.
                    if (da == null && db == null) return Integer.compare(a, b);
                    if (da == null) return 1;
                    if (db == null) return -1;
                    int hasil = da.compareTo(db);
                    if (hasil == 0) return Integer.compare(a, b);
                    return terbaru ? -hasil : hasil;
                });
                break;

            case 3: // Nominal terbesar
                Collections.sort(indeks, (a, b) -> Double.compare(
                        transaksiList.get(b).nominal, transaksiList.get(a).nominal));
                break;

            case 4: // Nominal terkecil
                Collections.sort(indeks, (a, b) -> Double.compare(
                        transaksiList.get(a).nominal, transaksiList.get(b).nominal));
                break;

            default:
                break;
        }

        return indeks;
    }

    private void tampilkanDaftarTransaksi() {
        // Anak pertama adalah txtDaftarTransaksi (keterangan),
        // sisanya baris transaksi yang dibuat ulang setiap kali.
        int jumlahAnak = listTransaksi.getChildCount();
        if (jumlahAnak > 1) {
            listTransaksi.removeViews(1, jumlahAnak - 1);
        }

        if (transaksiList.isEmpty()) {
            txtDaftarTransaksi.setText("Belum ada transaksi.");
            return;
        }

        List<Integer> indeks = indeksTampil();

        if (indeks.isEmpty()) {
            txtDaftarTransaksi.setText("Tidak ada transaksi yang cocok dengan filter.");
            return;
        }

        String jumlah = indeks.size() == transaksiList.size()
                ? transaksiList.size() + " transaksi"
                : "Menampilkan " + indeks.size() + " dari " + transaksiList.size();

        txtDaftarTransaksi.setText(jumlah + "  •  ketuk untuk edit/hapus");

        TypedValue ripple = new TypedValue();
        getTheme().resolveAttribute(
                android.R.attr.selectableItemBackground, ripple, true);

        int padding = dp(8);

        for (int index : indeks) {
            Transaksi t = transaksiList.get(index);

            View garis = new View(this);
            garis.setBackgroundColor(Color.parseColor("#E8E2EF"));
            listTransaksi.addView(garis, new LinearLayout.LayoutParams(
                    LinearLayout.LayoutParams.MATCH_PARENT, dp(1)));

            String bon = t.foto.isEmpty() ? "tanpa bon" : "bon: " + t.foto.size() + " foto";

            TextView baris = new TextView(this);
            baris.setText((index + 1) + ". " + t.jenis + " - " + t.kategori + "\n"
                    + t.tanggal + " | " + t.uraian + "\n"
                    + formatRupiah(t.nominal) + "   •   " + bon);
            baris.setTextColor(Color.parseColor("#6640A4"));
            baris.setTextSize(TypedValue.COMPLEX_UNIT_SP, 13);
            baris.setLineSpacing(dp(3), 1f);
            baris.setPadding(padding, padding, padding, padding);
            baris.setBackgroundResource(ripple.resourceId);
            baris.setClickable(true);
            baris.setFocusable(true);
            baris.setOnClickListener(v -> tampilkanMenuTransaksi(index));

            // Tandai baris yang sedang diedit.
            if (index == editIndex) {
                baris.setBackgroundColor(Color.parseColor("#F0EAFB"));
                baris.setTypeface(baris.getTypeface(), Typeface.BOLD);
            }

            listTransaksi.addView(baris, new LinearLayout.LayoutParams(
                    LinearLayout.LayoutParams.MATCH_PARENT,
                    LinearLayout.LayoutParams.WRAP_CONTENT));
        }
    }

    // =====================================================================
    // Buka ulang file Excel
    // =====================================================================

    /** Minta konfirmasi jika data di layar akan diganti. */
    private void konfirmasiGantiData(Runnable lanjut) {
        if (transaksiList.isEmpty() && fotoUmum.isEmpty()) {
            lanjut.run();
            return;
        }

        new AlertDialog.Builder(this)
                .setTitle("Ganti data di layar?")
                .setMessage("Transaksi dan foto yang sedang terbuka akan diganti dengan isi file. "
                        + "Ekspor dulu jika data saat ini belum disimpan ke Excel.")
                .setPositiveButton("Lanjutkan", (d, w) -> lanjut.run())
                .setNegativeButton("Batal", null)
                .show();
    }

    private void muatFileExcel(Uri uri) {
        findViewById(R.id.btnReopenExcel).setEnabled(false);
        txtStatus.setText("Sedang memuat data reimbursement...");

        executor.execute(() -> {
            File salinan = new File(getCacheDir(), "reopen_tmp.xlsx");

            try {
                // Salin dulu ke file sementara, lalu buka SEKALI dari disk.
                try (InputStream input = getContentResolver().openInputStream(uri)) {
                    if (input == null) {
                        throw new Exception("File Excel tidak dapat dibuka.");
                    }
                    salinKeFile(input, salinan);
                }

                ExcelHelper.ReimburseData data = ExcelHelper.readReimbursement(salinan);

                List<Transaksi> material = new ArrayList<>();
                List<Transaksi> jasa = new ArrayList<>();
                List<Transaksi> operasional = new ArrayList<>();

                tambahkanTransaksiDariBaris(material, data.material, "Material");
                tambahkanTransaksiDariBaris(jasa, data.jasa, "Jasa");
                tambahkanTransaksiDariBaris(operasional, data.operasional, "Operasional");

                // Foto: kompres ulang, lalu tempelkan kembali ke transaksinya
                // berdasarkan keterangan "Bon <Jenis> no. <n>".
                List<String> umumBaru = new ArrayList<>();
                int fotoGagal = 0;
                File sementara = new File(getCacheDir(), "reopen_foto.tmp");

                for (ExcelHelper.Gambar g : data.photos) {
                    String nama;

                    try {
                        tulisFile(sementara, g.data);
                        nama = FotoHelper.simpanDariFile(sementara, folderFoto());
                    } catch (Throwable t) {
                        fotoGagal++;
                        Log.w(TAG, "Foto dari Excel dilewati", t);
                        continue;
                    } finally {
                        sementara.delete();
                    }

                    Transaksi pemilik = null;
                    Matcher m = POLA_KETERANGAN_BON.matcher(g.keterangan);

                    if (m.find()) {
                        List<Transaksi> kelompok = "Material".equals(m.group(1)) ? material
                                : "Jasa".equals(m.group(1)) ? jasa
                                : operasional;
                        int nomor = Integer.parseInt(m.group(2));
                        if (nomor >= 1 && nomor <= kelompok.size()) {
                            pemilik = kelompok.get(nomor - 1);
                        }
                    }

                    if (pemilik != null) {
                        pemilik.foto.add(nama);
                    } else {
                        umumBaru.add(nama);
                    }
                }

                // Lepas referensi foto agar memori bisa dibebaskan.
                data.photos.clear();

                final List<Transaksi> transaksiBaru = new ArrayList<>();
                transaksiBaru.addAll(material);
                transaksiBaru.addAll(jasa);
                transaksiBaru.addAll(operasional);

                final Map<String, String> resume = data.resume;
                final double kas = parseAngka(resume.get("H20"));
                final int gagal = fotoGagal;

                runOnUiThread(() -> {
                    if (isFinishing() || isDestroyed()) return;

                    try {
                        keluarModeEdit(true);
                        kosongkanFormTransaksi();

                        transaksiList.clear();
                        transaksiList.addAll(transaksiBaru);

                        fotoUmum.clear();
                        fotoUmum.addAll(umumBaru);

                        editProject.setText(teksAman(resume.get("G10")));
                        editReimburseDate.setText(teksAman(resume.get("G11")));
                        editPIC.setText(teksAman(resume.get("G12")));
                        editKas.setText(kas > 0 ? angkaUntukInput(kas) : "");

                        perbaruiJumlahFoto();
                        perbaruiRekap();
                        simpanDraf();

                        int fotoTransaksi = 0;
                        for (Transaksi t : transaksiList) fotoTransaksi += t.foto.size();

                        String pesan = "File berhasil dimuat: "
                                + transaksiList.size() + " transaksi, "
                                + fotoTransaksi + " foto bon, "
                                + fotoUmum.size() + " lampiran lain.";
                        if (gagal > 0) pesan += "\n" + gagal + " gambar tidak bisa dibaca.";

                        txtStatus.setText(pesan);

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

            hasil.add(new Transaksi(jenis, tanggal, kategori, uraian, nominal, null));
        }
    }

    // =====================================================================
    // Ekspor Excel
    // =====================================================================

    private void eksporExcel() {
        String proyek = editProject.getText().toString().trim();
        String pic = editPIC.getText().toString().trim();
        String tanggalReimburse = editReimburseDate.getText().toString().trim();

        if (proyek.isEmpty() || pic.isEmpty() || tanggalReimburse.isEmpty()) {
            // Kolom ini ada di tab Input, jadi pindahkan dulu agar terlihat.
            pindahHalaman(R.id.nav_input);

            if (proyek.isEmpty()) {
                editProject.setError("Nama proyek wajib diisi");
                editProject.requestFocus();
            } else if (pic.isEmpty()) {
                editPIC.setError("PIC wajib diisi");
                editPIC.requestFocus();
            } else {
                editReimburseDate.setError("Tanggal reimburse wajib diisi");
            }

            Toast.makeText(this, "Lengkapi data aktivitas terlebih dahulu",
                    Toast.LENGTH_LONG).show();
            return;
        }

        if (editIndex >= 0) {
            pindahHalaman(R.id.nav_input);
            Toast.makeText(this,
                    "Simpan atau batalkan edit transaksi terlebih dahulu",
                    Toast.LENGTH_LONG).show();
            return;
        }

        // Periksa kapasitas (misalnya data dari file lama).
        for (String jenis : new String[]{"Material", "Jasa", "Operasional"}) {
            int jumlah = jumlahJenis(jenis, -1);
            if (jumlah > kapasitas(jenis)) {
                new AlertDialog.Builder(this)
                        .setTitle("Transaksi " + jenis + " terlalu banyak")
                        .setMessage("Ada " + jumlah + " transaksi " + jenis
                                + ", sedangkan template hanya menyediakan "
                                + kapasitas(jenis) + " baris.")
                        .setPositiveButton("Mengerti", null)
                        .show();
                return;
            }
        }

        if (!fotoForm.isEmpty()) {
            Toast.makeText(this,
                    "Ada foto bon di form yang belum disimpan ke transaksi",
                    Toast.LENGTH_LONG).show();
        }

        simpanDraf();

        // Snapshot agar data yang diekspor konsisten.
        final String proyekFinal = proyek;
        final String picFinal = pic;
        final String tanggalFinal = tanggalReimburse;
        final double kasFinal = ambilKas();
        final List<Transaksi> transaksiFinal = new ArrayList<>();
        for (Transaksi t : transaksiList) transaksiFinal.add(t.salin());
        final List<String> fotoUmumFinal = new ArrayList<>(fotoUmum);
        final String perusahaanFinal = pengaturan.namaPerusahaan();

        findViewById(R.id.btnEksporExcel).setEnabled(false);
        txtStatus.setText("Sedang membuat file Excel...");

        executor.execute(() -> {
            Uri outputUri = null;

            try {
                // Susun foto dengan urutan yang sama seperti penomoran Rincian.
                List<ExcelHelper.Gambar> gambar = new ArrayList<>();
                Map<Transaksi, String> keterangan = new IdentityHashMap<>();
                int nomorLampiran = 1;
                int fotoGagal = 0;

                for (String jenis : new String[]{"Material", "Jasa", "Operasional"}) {
                    int nomorTransaksi = 1;

                    for (Transaksi t : transaksiFinal) {
                        if (!t.jenis.equals(jenis)) continue;

                        List<String> nomor = new ArrayList<>();

                        for (String namaFoto : t.foto) {
                            try {
                                File f = fileFoto(namaFoto);
                                int[] ukuran = FotoHelper.ukuran(f);
                                gambar.add(new ExcelHelper.Gambar(
                                        bacaFile(f), ukuran[0], ukuran[1],
                                        "Bon " + jenis + " no. " + nomorTransaksi
                                                + " | " + t.uraian
                                                + " | " + formatRupiah(t.nominal)));
                                nomor.add(String.valueOf(nomorLampiran++));
                            } catch (Throwable e) {
                                fotoGagal++;
                                Log.w(TAG, "Foto bon tidak terbaca: " + namaFoto, e);
                            }
                        }

                        keterangan.put(t, nomor.isEmpty()
                                ? "Tanpa bon"
                                : "Bon: lampiran " + gabung(nomor));

                        nomorTransaksi++;
                    }
                }

                for (String namaFoto : fotoUmumFinal) {
                    try {
                        File f = fileFoto(namaFoto);
                        int[] ukuran = FotoHelper.ukuran(f);
                        gambar.add(new ExcelHelper.Gambar(
                                bacaFile(f), ukuran[0], ukuran[1], "Lampiran lain"));
                        nomorLampiran++;
                    } catch (Throwable e) {
                        fotoGagal++;
                        Log.w(TAG, "Lampiran lain tidak terbaca: " + namaFoto, e);
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

                // Hanya isi sel input. Sel total & saldo (H15-H18, H22)
                // dibiarkan berupa rumus template.
                Map<String, Object> resume = new LinkedHashMap<>();
                resume.put("G10", proyekFinal);
                resume.put("G11", tanggalFinal);
                resume.put("G12", picFinal);
                resume.put("H20", kasFinal);

                Map<String, Object> rincian = buatDataRincian(transaksiFinal, keterangan);

                // Judul dokumen mengikuti nama perusahaan di Pengaturan.
                if (!perusahaanFinal.isEmpty()) {
                    resume.put("E3", perusahaanFinal + "\nREIMBURSEMENT");
                    rincian.put("A2", perusahaanFinal + "\nRINCIAN PENGELUARAN");
                }

                allData.put("Resume", resume);
                allData.put("Rincian", rincian);

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

                    ExcelHelper.writeWorkbook(template, output, allData, gambar);
                }

                gambar.clear();

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

                // FITUR 8: catat ke riwayat ekspor.
                double total = 0;
                for (Transaksi t : transaksiFinal) total += t.nominal;

                riwayat.tambah(new RiwayatEkspor.Item(
                        fileName, outputUri.toString(), proyekFinal, tanggalFinal,
                        total, transaksiFinal.size(), System.currentTimeMillis()));

                final Uri hasilUri = outputUri;
                final int gagal = fotoGagal;

                runOnUiThread(() -> {
                    if (isFinishing() || isDestroyed()) return;

                    findViewById(R.id.btnEksporExcel).setEnabled(true);

                    terakhirEksporUri = hasilUri.toString();
                    terakhirEksporNama = fileName;
                    perbaruiTombolBagikan();
                    simpanDraf();

                    String pesan = "Ekspor berhasil!\n" + FOLDER_PATH + fileName;

                    if (gagal > 0) {
                        pesan += "\nPerhatian: " + gagal
                                + " foto tidak bisa dibaca dan tidak ikut terlampir.";
                    }

                    txtStatus.setText(pesan);

                    // FITUR 2: tawarkan langsung untuk dibagikan.
                    new AlertDialog.Builder(this)
                            .setTitle("Ekspor berhasil")
                            .setMessage("File disimpan di " + FOLDER_PATH + fileName
                                    + (gagal > 0 ? "\n\n" + gagal + " foto tidak ikut terlampir." : ""))
                            .setPositiveButton("Bagikan", (d, w) -> bagikanFile(hasilUri, fileName))
                            .setNegativeButton("Tutup", null)
                            .show();
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
     * A = nomor, B = tanggal, C = nominal, D = kategori, E = uraian,
     * F = keterangan bon.
     *
     * Material    : baris 9-37   (total di C38 tetap rumus template)
     * Jasa        : baris 43-49  (total di C50 tetap rumus template)
     * Operasional : baris 55-117 (total di C118 tetap rumus template)
     */
    private Map<String, Object> buatDataRincian(
            List<Transaksi> data, Map<Transaksi, String> keterangan) throws Exception {

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
                        throw new Exception("Transaksi Material maksimal 29 baris.");
                    }
                    row = materialRow++;
                    cells.put("A" + row, nomorMaterial++);
                    break;

                case "Jasa":
                    if (jasaRow > 49) {
                        throw new Exception("Transaksi Jasa maksimal 7 baris.");
                    }
                    row = jasaRow++;
                    cells.put("A" + row, nomorJasa++);
                    break;

                default:
                    if (operasionalRow > 117) {
                        throw new Exception("Transaksi Operasional maksimal 63 baris.");
                    }
                    row = operasionalRow++;
                    cells.put("A" + row, nomorOperasional++);
                    break;
            }

            cells.put("B" + row, t.tanggal);
            cells.put("C" + row, t.nominal);
            cells.put("D" + row, t.kategori);
            cells.put("E" + row, t.uraian);

            String ket = keterangan.get(t);
            if (ket != null) cells.put("F" + row, ket);
        }

        return cells;
    }

    // =====================================================================
    // FITUR 2 & 8: Bagikan dan riwayat ekspor
    // =====================================================================

    private void perbaruiTombolBagikan() {
        btnBagikan.setVisibility(terakhirEksporUri.isEmpty() ? View.GONE : View.VISIBLE);
        if (!terakhirEksporNama.isEmpty()) {
            btnBagikan.setText("Bagikan: " + terakhirEksporNama);
        }
    }

    private void bagikanFile(Uri uri, String nama) {
        try {
            Intent kirim = new Intent(Intent.ACTION_SEND);
            kirim.setType(MIME_XLSX);
            kirim.putExtra(Intent.EXTRA_STREAM, uri);
            kirim.putExtra(Intent.EXTRA_SUBJECT, nama);
            kirim.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION);

            startActivity(Intent.createChooser(kirim, "Bagikan file reimbursement"));
        } catch (Exception e) {
            Log.e(TAG, "Gagal membagikan", e);
            Toast.makeText(this, "Tidak ada aplikasi untuk membagikan file",
                    Toast.LENGTH_LONG).show();
        }
    }

    private void tampilkanRiwayat() {
        List<RiwayatEkspor.Item> daftar = riwayat.muat();

        if (daftar.isEmpty()) {
            Toast.makeText(this, "Belum ada riwayat ekspor", Toast.LENGTH_SHORT).show();
            return;
        }

        SimpleDateFormat formatWaktu =
                new SimpleDateFormat("dd MMM yyyy, HH:mm", localeIndonesia);

        String[] baris = new String[daftar.size()];

        for (int i = 0; i < daftar.size(); i++) {
            RiwayatEkspor.Item item = daftar.get(i);
            baris[i] = item.proyek + "\n"
                    + item.jumlahTransaksi + " transaksi  •  " + formatRupiah(item.total) + "\n"
                    + "Diekspor " + formatWaktu.format(new Date(item.waktu));
        }

        new AlertDialog.Builder(this)
                .setTitle("Riwayat Ekspor")
                .setItems(baris, (d, which) -> tampilkanMenuRiwayat(daftar.get(which)))
                .setNegativeButton("Tutup", null)
                .show();
    }

    private void tampilkanMenuRiwayat(RiwayatEkspor.Item item) {
        new AlertDialog.Builder(this)
                .setTitle(item.namaFile)
                .setItems(new String[]{"Buka di aplikasi", "Bagikan", "Hapus dari riwayat"},
                        (d, which) -> {
                            Uri uri = Uri.parse(item.uri);

                            if (which == 0) {
                                konfirmasiGantiData(() -> muatFileExcel(uri));
                            } else if (which == 1) {
                                bagikanFile(uri, item.namaFile);
                            } else {
                                riwayat.hapus(item.uri);
                                Toast.makeText(this,
                                        "Dihapus dari riwayat (file Excel tetap ada)",
                                        Toast.LENGTH_SHORT).show();
                            }
                        })
                .setNegativeButton("Tutup", null)
                .show();
    }

    // =====================================================================
    // Draf otomatis (tidak hilang saat layar diputar / app ditutup)
    // =====================================================================

    private void simpanDraf() {
        try {
            JSONObject root = new JSONObject();
            root.put("versi", 2);
            root.put("proyek", editProject.getText().toString());
            root.put("pic", editPIC.getText().toString());
            root.put("tanggalReimburse", editReimburseDate.getText().toString());
            root.put("kas", editKas.getText().toString());
            root.put("jenis", jenisTerpilih);
            root.put("terakhirEksporUri", terakhirEksporUri);
            root.put("terakhirEksporNama", terakhirEksporNama);

            JSONArray daftar = new JSONArray();
            for (Transaksi t : transaksiList) {
                JSONObject o = new JSONObject();
                o.put("jenis", t.jenis);
                o.put("tanggal", t.tanggal);
                o.put("kategori", t.kategori);
                o.put("uraian", t.uraian);
                o.put("nominal", t.nominal);
                o.put("foto", keJson(t.foto));
                daftar.put(o);
            }
            root.put("transaksi", daftar);

            root.put("fotoUmum", keJson(fotoUmum));
            root.put("fotoForm", keJson(fotoForm));

            byte[] isi = root.toString().getBytes(StandardCharsets.UTF_8);

            // Tulis ke file sementara dulu, baru diganti, agar draf
            // tidak rusak jika aplikasi tertutup saat menyimpan.
            File tujuan = new File(getFilesDir(), DRAFT_FILE);
            File sementara = new File(getFilesDir(), DRAFT_FILE + ".tmp");
            tulisFile(sementara, isi);

            if (!sementara.renameTo(tujuan)) {
                tulisFile(tujuan, isi);
                sementara.delete();
            }
        } catch (Exception e) {
            Log.w(TAG, "Gagal menyimpan draf", e);
        }
    }

    private void muatDraf() {
        File file = new File(getFilesDir(), DRAFT_FILE);

        if (!file.exists()) {
            // Pertama kali dipakai: isi dari pengaturan.
            editKas.setText(angkaUntukInput(pengaturan.kasDefault()));
            editPIC.setText(pengaturan.picDefault());
            return;
        }

        try {
            JSONObject root = new JSONObject(
                    new String(bacaFile(file), StandardCharsets.UTF_8));

            editProject.setText(root.optString("proyek"));
            editPIC.setText(root.optString("pic"));
            editReimburseDate.setText(root.optString("tanggalReimburse"));

            double kas = parseAngka(root.optString("kas"));
            editKas.setText(kas > 0 ? angkaUntukInput(kas) : "");

            jenisTerpilih = root.optString("jenis", "Material");
            perbaruiKategori(null);

            terakhirEksporUri = root.optString("terakhirEksporUri", "");
            terakhirEksporNama = root.optString("terakhirEksporNama", "");

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
                            o.optDouble("nominal", 0),
                            dariJson(o.optJSONArray("foto"))));
                }
            }

            fotoUmum.clear();
            fotoUmum.addAll(dariJson(root.optJSONArray("fotoUmum")));

            fotoForm.clear();
            fotoForm.addAll(dariJson(root.optJSONArray("fotoForm")));

            migrasiFotoVersiLama(root);
        } catch (Exception e) {
            Log.w(TAG, "Draf rusak, mulai dari kosong", e);
            editKas.setText(angkaUntukInput(pengaturan.kasDefault()));
        }
    }

    /**
     * Draf dari versi sebelumnya menyimpan foto sebagai URI galeri
     * ("fotoUris") dan file di folder lama ("fotoLama"). Keduanya
     * dikompres dan dipindahkan ke lampiran lain.
     */
    private void migrasiFotoVersiLama(JSONObject root) throws Exception {
        final List<String> uris = dariJson(root.optJSONArray("fotoUris"));
        final List<String> lama = dariJson(root.optJSONArray("fotoLama"));

        if (uris.isEmpty() && lama.isEmpty()) return;

        final File folderLama = new File(getFilesDir(), FOLDER_FOTO_LAMA);

        executor.execute(() -> {
            List<String> hasil = new ArrayList<>();

            for (String nama : lama) {
                try {
                    hasil.add(FotoHelper.kompresFile(new File(folderLama, nama), folderFoto()));
                } catch (Throwable t) {
                    Log.w(TAG, "Migrasi foto lama dilewati: " + nama, t);
                }
            }

            for (String uri : uris) {
                try (InputStream input = getContentResolver().openInputStream(Uri.parse(uri))) {
                    if (input == null) continue;
                    hasil.add(FotoHelper.simpanTerkompres(input, getCacheDir(), folderFoto()));
                } catch (Throwable t) {
                    Log.w(TAG, "Migrasi foto galeri dilewati: " + uri, t);
                }
            }

            File[] sisa = folderLama.listFiles();
            if (sisa != null) for (File f : sisa) f.delete();
            folderLama.delete();

            runOnUiThread(() -> {
                if (isFinishing() || isDestroyed()) return;
                fotoUmum.addAll(hasil);
                perbaruiJumlahFoto();
                simpanDraf();
            });
        });
    }

    private void konfirmasiMulaiBaru() {
        new AlertDialog.Builder(this)
                .setTitle("Mulai reimbursement baru?")
                .setMessage("Semua data di layar (proyek, transaksi, dan foto bon) "
                        + "akan dikosongkan. File Excel yang sudah diekspor dan "
                        + "riwayat ekspor tidak terpengaruh.")
                .setPositiveButton("Kosongkan", (d, w) -> mulaiBaru())
                .setNegativeButton("Batal", null)
                .show();
    }

    private void mulaiBaru() {
        keluarModeEdit(true);
        kosongkanFormTransaksi();

        editProject.setText("");
        editPIC.setText(pengaturan.picDefault());
        editReimburseDate.setText("");
        editKas.setText(angkaUntukInput(pengaturan.kasDefault()));

        transaksiList.clear();
        fotoUmum.clear();
        pilihJenis("Material");

        perbaruiJumlahFoto();
        perbaruiRekap();
        simpanDraf();

        // Semua foto sekarang tidak terpakai, jadi dihapus.
        bersihkanFotoTakTerpakai();

        txtStatus.setText("Siap digunakan.");
    }

    // =====================================================================
    // FITUR 5: Format ribuan otomatis
    // =====================================================================

    /** Menambahkan titik ribuan saat mengetik, mis. 1500000 -> 1.500.000. */
    private static class FormatRibuan implements TextWatcher {
        private final EditText target;
        private final Runnable setelahBerubah;
        private boolean sibuk = false;

        FormatRibuan(EditText target, Runnable setelahBerubah) {
            this.target = target;
            this.setelahBerubah = setelahBerubah;
        }

        @Override
        public void beforeTextChanged(CharSequence s, int start, int count, int after) {
        }

        @Override
        public void onTextChanged(CharSequence s, int start, int before, int count) {
        }

        @Override
        public void afterTextChanged(Editable s) {
            if (sibuk) return;

            String asli = s.toString();
            String baru = formatRibuanDariTeks(asli);

            if (!baru.equals(asli)) {
                int kursor = target.getSelectionEnd();
                if (kursor < 0 || kursor > asli.length()) kursor = asli.length();

                // Jaga posisi kursor relatif terhadap digit di sebelah kanannya.
                int digitDiKanan = 0;
                for (int i = kursor; i < asli.length(); i++) {
                    if (Character.isDigit(asli.charAt(i))) digitDiKanan++;
                }

                sibuk = true;
                target.setText(baru);

                int posisi = baru.length();
                int dihitung = 0;
                while (posisi > 0 && dihitung < digitDiKanan) {
                    posisi--;
                    if (Character.isDigit(baru.charAt(posisi))) dihitung++;
                }

                target.setSelection(posisi);
                sibuk = false;
            }

            if (setelahBerubah != null) setelahBerubah.run();
        }
    }

    /** "1500000" atau "1.5000.00" -> "1.500.000". */
    static String formatRibuanDariTeks(String teks) {
        String digit = teks.replaceAll("[^0-9]", "").replaceFirst("^0+(?=\\d)", "");

        if (digit.length() > 15) digit = digit.substring(0, 15);
        if (digit.isEmpty()) return "";

        StringBuilder hasil = new StringBuilder();
        int awal = digit.length() % 3;
        if (awal == 0) awal = 3;

        hasil.append(digit, 0, awal);
        for (int i = awal; i < digit.length(); i += 3) {
            hasil.append('.').append(digit, i, i + 3);
        }

        return hasil.toString();
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

    private double ambilKas() {
        double kas = parseAngka(editKas.getText().toString());
        return kas > 0 ? kas : 0;
    }

    /**
     * Mengubah teks angka menjadi double.
     * Contoh: "150000", "150.000", "150000,5", "Rp150.000", " 3,000,000 ".
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

    /** Angka untuk kolom input, dengan titik ribuan (dibulatkan ke rupiah). */
    private static String angkaUntukInput(double angka) {
        return formatRibuanDariTeks(String.valueOf(Math.round(angka)));
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

    private static String teksAman(String s) {
        return s == null ? "" : s;
    }

    private static String gabung(List<String> daftar) {
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < daftar.size(); i++) {
            if (i > 0) sb.append(", ");
            sb.append(daftar.get(i));
        }
        return sb.toString();
    }

    private static JSONArray keJson(List<String> daftar) {
        JSONArray array = new JSONArray();
        for (String s : daftar) array.put(s);
        return array;
    }

    private static List<String> dariJson(JSONArray array) throws Exception {
        List<String> hasil = new ArrayList<>();
        if (array == null) return hasil;
        for (int i = 0; i < array.length(); i++) hasil.add(array.getString(i));
        return hasil;
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

    private static byte[] bacaFile(File file) throws Exception {
        try (InputStream input = new FileInputStream(file);
             ByteArrayOutputStream output = new ByteArrayOutputStream()) {
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

    // =====================================================================
    // Model data
    // =====================================================================

    private static class Transaksi {
        final String jenis;
        final String tanggal;
        final String kategori;
        final String uraian;
        final double nominal;

        /** Nama file foto bon (boleh kosong: transaksi tanpa bon). */
        final List<String> foto = new ArrayList<>();

        Transaksi(String jenis, String tanggal, String kategori,
                  String uraian, double nominal, List<String> foto) {
            this.jenis = jenis;
            this.tanggal = tanggal;
            this.kategori = kategori;
            this.uraian = uraian;
            this.nominal = nominal;
            if (foto != null) this.foto.addAll(foto);
        }

        Transaksi salin() {
            return new Transaksi(jenis, tanggal, kategori, uraian, nominal, foto);
        }
    }
}
