package com.etdah.rembuseapps;

import android.graphics.Bitmap;
import android.graphics.BitmapFactory;
import android.graphics.Matrix;
import android.media.ExifInterface;

import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.UUID;

/**
 * Mengompres foto bon agar file Excel tetap kecil.
 *
 * Foto HP biasanya 3-5 MB. Setelah diperkecil ke sisi terpanjang
 * 1600 piksel dengan kualitas JPEG 80, ukurannya sekitar 200-400 KB
 * dan tulisan di bon masih jelas terbaca.
 */
public final class FotoHelper {

    /** Sisi terpanjang foto setelah dikompres (piksel). */
    private static final int SISI_MAKS = 1600;

    /** Kualitas JPEG (0-100). */
    private static final int KUALITAS = 80;

    private FotoHelper() {
    }

    /**
     * Membaca foto dari stream, memperbaiki rotasi, mengompres,
     * lalu menyimpannya ke folderTujuan.
     *
     * @return nama file hasil (bukan path lengkap).
     */
    public static String simpanTerkompres(
            InputStream input,
            File folderSementara,
            File folderTujuan
    ) throws Exception {

        if (!folderSementara.exists()) folderSementara.mkdirs();
        if (!folderTujuan.exists()) folderTujuan.mkdirs();

        File sementara = new File(folderSementara,
                "foto_masuk_" + UUID.randomUUID() + ".tmp");

        try {
            try (OutputStream output = new FileOutputStream(sementara)) {
                byte[] buffer = new byte[8192];
                int jumlah;
                while ((jumlah = input.read(buffer)) != -1) {
                    output.write(buffer, 0, jumlah);
                }
            }

            return kompresFile(sementara, folderTujuan);
        } finally {
            sementara.delete();
        }
    }

    /** Sama seperti simpanTerkompres, tetapi sumbernya file di disk. */
    public static String kompresFile(File sumber, File folderTujuan) throws Exception {
        if (!folderTujuan.exists()) folderTujuan.mkdirs();

        // 1. Baca ukuran asli tanpa memuat gambar ke memori.
        BitmapFactory.Options batas = new BitmapFactory.Options();
        batas.inJustDecodeBounds = true;
        BitmapFactory.decodeFile(sumber.getAbsolutePath(), batas);

        if (batas.outWidth <= 0 || batas.outHeight <= 0) {
            throw new Exception("File bukan gambar yang didukung.");
        }

        // 2. Muat versi yang sudah diperkecil (hemat memori).
        int sample = 1;
        int sisiTerpanjang = Math.max(batas.outWidth, batas.outHeight);
        while (sisiTerpanjang / (sample * 2) >= SISI_MAKS) {
            sample *= 2;
        }

        BitmapFactory.Options opsi = new BitmapFactory.Options();
        opsi.inSampleSize = sample;

        Bitmap bitmap = BitmapFactory.decodeFile(sumber.getAbsolutePath(), opsi);

        if (bitmap == null) {
            throw new Exception("Gambar tidak dapat dibaca.");
        }

        // 3. Perkecil sampai pas, sekaligus putar sesuai data EXIF
        //    (foto dari kamera sering tersimpan miring).
        Matrix matrix = new Matrix();
        int rotasi = rotasiExif(sumber);
        if (rotasi != 0) matrix.postRotate(rotasi);

        float skala = Math.min(1f,
                (float) SISI_MAKS / Math.max(bitmap.getWidth(), bitmap.getHeight()));
        if (skala < 1f) matrix.postScale(skala, skala);

        if (rotasi != 0 || skala < 1f) {
            Bitmap hasil = Bitmap.createBitmap(
                    bitmap, 0, 0, bitmap.getWidth(), bitmap.getHeight(), matrix, true);
            if (hasil != bitmap) {
                bitmap.recycle();
                bitmap = hasil;
            }
        }

        // 4. Simpan sebagai JPEG.
        String nama = "bon_" + UUID.randomUUID() + ".jpg";
        File tujuan = new File(folderTujuan, nama);

        try (OutputStream output = new FileOutputStream(tujuan)) {
            if (!bitmap.compress(Bitmap.CompressFormat.JPEG, KUALITAS, output)) {
                throw new Exception("Gagal mengompres foto.");
            }
        } finally {
            bitmap.recycle();
        }

        return nama;
    }

    /**
     * Untuk foto yang dibaca dari file Excel: jika sudah cukup kecil,
     * disalin apa adanya (agar kualitas tidak turun karena dikompres
     * berulang kali). Jika masih besar, dikompres.
     */
    public static String simpanDariFile(File sumber, File folderTujuan) throws Exception {
        BitmapFactory.Options batas = new BitmapFactory.Options();
        batas.inJustDecodeBounds = true;
        BitmapFactory.decodeFile(sumber.getAbsolutePath(), batas);

        if (batas.outWidth <= 0 || batas.outHeight <= 0) {
            throw new Exception("File bukan gambar yang didukung.");
        }

        boolean sudahKecil = Math.max(batas.outWidth, batas.outHeight) <= SISI_MAKS
                && sumber.length() <= 800 * 1024;

        if (!sudahKecil) {
            return kompresFile(sumber, folderTujuan);
        }

        if (!folderTujuan.exists()) folderTujuan.mkdirs();

        String ekstensi = "image/png".equals(batas.outMimeType) ? ".png" : ".jpg";
        String nama = "bon_" + UUID.randomUUID() + ekstensi;

        try (InputStream input = new java.io.FileInputStream(sumber);
             OutputStream output = new FileOutputStream(new File(folderTujuan, nama))) {
            byte[] buffer = new byte[8192];
            int jumlah;
            while ((jumlah = input.read(buffer)) != -1) {
                output.write(buffer, 0, jumlah);
            }
        }

        return nama;
    }

    /** Lebar dan tinggi gambar tanpa memuatnya ke memori. */
    public static int[] ukuran(File file) {
        BitmapFactory.Options batas = new BitmapFactory.Options();
        batas.inJustDecodeBounds = true;
        BitmapFactory.decodeFile(file.getAbsolutePath(), batas);
        return new int[]{Math.max(batas.outWidth, 1), Math.max(batas.outHeight, 1)};
    }

    private static int rotasiExif(File file) {
        try {
            ExifInterface exif = new ExifInterface(file.getAbsolutePath());
            int orientasi = exif.getAttributeInt(
                    ExifInterface.TAG_ORIENTATION,
                    ExifInterface.ORIENTATION_NORMAL);

            switch (orientasi) {
                case ExifInterface.ORIENTATION_ROTATE_90:
                    return 90;
                case ExifInterface.ORIENTATION_ROTATE_180:
                    return 180;
                case ExifInterface.ORIENTATION_ROTATE_270:
                    return 270;
                default:
                    return 0;
            }
        } catch (Throwable t) {
            return 0;
        }
    }
}
