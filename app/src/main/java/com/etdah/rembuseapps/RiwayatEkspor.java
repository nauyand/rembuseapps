package com.etdah.rembuseapps;

import org.json.JSONArray;
import org.json.JSONObject;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;

/** Menyimpan daftar file Excel yang pernah diekspor (maksimal 50 terakhir). */
public final class RiwayatEkspor {

    private static final String NAMA_FILE = "riwayat_ekspor.json";
    private static final int MAKS = 50;

    public static class Item {
        public final String namaFile;
        public final String uri;
        public final String proyek;
        public final String tanggal;
        public final double total;
        public final int jumlahTransaksi;
        public final long waktu;

        public Item(String namaFile, String uri, String proyek, String tanggal,
                    double total, int jumlahTransaksi, long waktu) {
            this.namaFile = namaFile;
            this.uri = uri;
            this.proyek = proyek;
            this.tanggal = tanggal;
            this.total = total;
            this.jumlahTransaksi = jumlahTransaksi;
            this.waktu = waktu;
        }
    }

    private final File file;

    public RiwayatEkspor(File folder) {
        this.file = new File(folder, NAMA_FILE);
    }

    /** Daftar riwayat, terbaru di atas. */
    public synchronized List<Item> muat() {
        List<Item> hasil = new ArrayList<>();

        if (!file.exists()) return hasil;

        try (InputStream input = new FileInputStream(file)) {
            byte[] isi = new byte[(int) file.length()];
            int dibaca = 0;
            while (dibaca < isi.length) {
                int n = input.read(isi, dibaca, isi.length - dibaca);
                if (n < 0) break;
                dibaca += n;
            }

            JSONArray daftar = new JSONArray(new String(isi, 0, dibaca, StandardCharsets.UTF_8));

            for (int i = 0; i < daftar.length(); i++) {
                JSONObject o = daftar.getJSONObject(i);
                hasil.add(new Item(
                        o.optString("namaFile"),
                        o.optString("uri"),
                        o.optString("proyek"),
                        o.optString("tanggal"),
                        o.optDouble("total", 0),
                        o.optInt("jumlahTransaksi", 0),
                        o.optLong("waktu", 0)));
            }
        } catch (Exception e) {
            android.util.Log.w("REMBUSE", "Riwayat tidak terbaca", e);
        }

        return hasil;
    }

    public synchronized void tambah(Item item) {
        List<Item> daftar = muat();
        daftar.add(0, item);

        while (daftar.size() > MAKS) {
            daftar.remove(daftar.size() - 1);
        }

        simpan(daftar);
    }

    public synchronized void hapus(String uri) {
        List<Item> daftar = muat();
        List<Item> sisa = new ArrayList<>();

        for (Item item : daftar) {
            if (!item.uri.equals(uri)) sisa.add(item);
        }

        simpan(sisa);
    }

    private void simpan(List<Item> daftar) {
        try {
            JSONArray array = new JSONArray();

            for (Item item : daftar) {
                JSONObject o = new JSONObject();
                o.put("namaFile", item.namaFile);
                o.put("uri", item.uri);
                o.put("proyek", item.proyek);
                o.put("tanggal", item.tanggal);
                o.put("total", item.total);
                o.put("jumlahTransaksi", item.jumlahTransaksi);
                o.put("waktu", item.waktu);
                array.put(o);
            }

            try (OutputStream output = new FileOutputStream(file)) {
                output.write(array.toString().getBytes(StandardCharsets.UTF_8));
            }
        } catch (Exception e) {
            android.util.Log.w("REMBUSE", "Riwayat gagal disimpan", e);
        }
    }
}
