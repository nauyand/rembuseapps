package com.etdah.rembuseapps;

import android.content.Context;
import android.content.SharedPreferences;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

/**
 * Pengaturan aplikasi yang bisa diubah pengguna tanpa mengubah kode:
 * nama perusahaan, PIC default, kas default, dan daftar kategori.
 */
public final class Pengaturan {

    private static final String NAMA_PREFS = "pengaturan";

    private static final String KEY_PERUSAHAAN = "namaPerusahaan";
    private static final String KEY_PIC = "picDefault";
    private static final String KEY_KAS = "kasDefault";
    private static final String KEY_KATEGORI = "kategori_";

    public static final String PERUSAHAAN_BAWAAN = "PT AXIUS TEKNOLOGI SILIWANGI";
    public static final double KAS_BAWAAN = 3000000.0;

    private final SharedPreferences prefs;

    public Pengaturan(Context context) {
        prefs = context.getSharedPreferences(NAMA_PREFS, Context.MODE_PRIVATE);
    }

    public String namaPerusahaan() {
        return prefs.getString(KEY_PERUSAHAAN, PERUSAHAAN_BAWAAN).trim();
    }

    public String picDefault() {
        return prefs.getString(KEY_PIC, "").trim();
    }

    public double kasDefault() {
        return Double.longBitsToDouble(
                prefs.getLong(KEY_KAS, Double.doubleToLongBits(KAS_BAWAAN)));
    }

    /** Daftar kategori untuk jenis tertentu (tidak pernah kosong). */
    public List<String> kategori(String jenis) {
        String tersimpan = prefs.getString(KEY_KATEGORI + jenis, null);
        List<String> hasil = tersimpan == null
                ? kategoriBawaan(jenis)
                : pisahkan(tersimpan);
        return hasil.isEmpty() ? kategoriBawaan(jenis) : hasil;
    }

    /** Kategori dalam bentuk teks dipisah koma, untuk ditampilkan di form. */
    public String kategoriTeks(String jenis) {
        return gabung(kategori(jenis));
    }

    public void simpan(String namaPerusahaan, String picDefault, double kasDefault,
                       String kategoriMaterial, String kategoriJasa,
                       String kategoriOperasional) {
        prefs.edit()
                .putString(KEY_PERUSAHAAN, namaPerusahaan.trim())
                .putString(KEY_PIC, picDefault.trim())
                .putLong(KEY_KAS, Double.doubleToLongBits(Math.max(kasDefault, 0)))
                .putString(KEY_KATEGORI + "Material", gabung(pisahkan(kategoriMaterial)))
                .putString(KEY_KATEGORI + "Jasa", gabung(pisahkan(kategoriJasa)))
                .putString(KEY_KATEGORI + "Operasional", gabung(pisahkan(kategoriOperasional)))
                .apply();
    }

    public void kembalikanBawaan() {
        prefs.edit().clear().apply();
    }

    public static List<String> kategoriBawaan(String jenis) {
        switch (jenis) {
            case "Material":
                return new ArrayList<>(Arrays.asList("Utama", "Pendukung"));
            case "Jasa":
                return new ArrayList<>(Arrays.asList("Axius", "Subcon"));
            default:
                return new ArrayList<>(Arrays.asList(
                        "Transportasi", "Meal", "Akomodasi",
                        "ATK", "Entertainment", "Toolkit"));
        }
    }

    /** "Utama, Pendukung,,  Lain " -> [Utama, Pendukung, Lain] (tanpa duplikat). */
    static List<String> pisahkan(String teks) {
        List<String> hasil = new ArrayList<>();
        if (teks == null) return hasil;

        for (String bagian : teks.split("[,\\n]")) {
            String s = bagian.trim();
            if (!s.isEmpty() && !hasil.contains(s)) hasil.add(s);
        }
        return hasil;
    }

    static String gabung(List<String> daftar) {
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < daftar.size(); i++) {
            if (i > 0) sb.append(", ");
            sb.append(daftar.get(i));
        }
        return sb.toString();
    }
}
