package com.etdah.rembuseapps;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFShape;

import java.io.File;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;

public class ExcelHelper {

    /** Ukuran 1 piksel dalam satuan EMU (dipakai posisi gambar di Excel). */
    private static final int EMU_PER_PIKSEL = 9525;

    /** Tinggi foto di sheet Lampiran (piksel). */
    private static final float TINGGI_FOTO = 420f;

    /** Lebar maksimal foto di sheet Lampiran (piksel). */
    private static final float LEBAR_MAKS_FOTO = 700f;

    /** Satu gambar beserta keterangannya. */
    public static class Gambar {
        public final byte[] data;
        public final int lebar;
        public final int tinggi;
        public final String keterangan;

        public Gambar(byte[] data, int lebar, int tinggi, String keterangan) {
            this.data = data;
            this.lebar = lebar;
            this.tinggi = tinggi;
            this.keterangan = keterangan == null ? "" : keterangan;
        }
    }

    /** Semua data yang dibaca dari file reimbursement. */
    public static class ReimburseData {
        public final Map<String, String> resume = new LinkedHashMap<>();
        public final List<List<String>> material = new ArrayList<>();
        public final List<List<String>> jasa = new ArrayList<>();
        public final List<List<String>> operasional = new ArrayList<>();
        public final List<Gambar> photos = new ArrayList<>();
    }

    // =====================================================================
    // Membaca
    // =====================================================================

    /**
     * Mengambil teks sel. Sengaja TIDAK memakai DataFormatter: pada format
     * tertentu (misalnya format akuntansi) DataFormatter memanggil
     * java.awt.Color yang tidak ada di Android sehingga aplikasi crash.
     * Tanggal selalu dikembalikan sebagai dd/MM/yyyy.
     */
    private static String cellText(Cell cell) {
        if (cell == null) return "";

        try {
            CellType type = cell.getCellType();

            // Untuk sel rumus, gunakan hasil terakhir yang tersimpan.
            if (type == CellType.FORMULA) {
                type = cell.getCachedFormulaResultType();
            }

            switch (type) {
                case STRING:
                    return cell.getStringCellValue().trim();

                case NUMERIC:
                    if (DateUtil.isCellDateFormatted(cell)) {
                        return new SimpleDateFormat("dd/MM/yyyy", Locale.US)
                                .format(cell.getDateCellValue());
                    }
                    return angkaPolos(cell.getNumericCellValue());

                case BOOLEAN:
                    return String.valueOf(cell.getBooleanCellValue());

                default:
                    return "";
            }
        } catch (Throwable t) {
            // Jangan sampai satu sel aneh membuat seluruh file gagal dibuka.
            return "";
        }
    }

    private static String angkaPolos(double angka) {
        if (angka == Math.rint(angka) && Math.abs(angka) < 1e15) {
            return String.valueOf((long) angka);
        }
        return java.math.BigDecimal.valueOf(angka)
                .stripTrailingZeros()
                .toPlainString();
    }

    private static String teksSel(Sheet sheet, int baris, int kolom) {
        if (baris < 0) return "";
        Row row = sheet.getRow(baris);
        return row == null ? "" : cellText(row.getCell(kolom));
    }

    /** Membaca satu sel (dipakai untuk verifikasi hasil ekspor). */
    public static String readCell(
            InputStream input,
            String sheetName,
            String cellAddress
    ) throws Exception {
        try (Workbook workbook = WorkbookFactory.create(input)) {
            Sheet sheet = workbook.getSheet(sheetName);

            if (sheet == null) {
                throw new IllegalArgumentException(
                        "Sheet tidak ditemukan: " + sheetName);
            }

            CellAddress address = new CellAddress(cellAddress);
            return teksSel(sheet, address.getRow(), address.getColumn());
        }
    }

    /**
     * Membaca seluruh isi file reimbursement dengan SEKALI buka.
     * File dibuka langsung dari disk dalam mode baca-saja agar hemat memori.
     */
    public static ReimburseData readReimbursement(File file) throws Exception {
        ReimburseData data = new ReimburseData();

        try (Workbook workbook = WorkbookFactory.create(file, null, true)) {
            Sheet resume = workbook.getSheet("Resume");
            Sheet rincian = workbook.getSheet("Rincian");

            if (resume == null || rincian == null) {
                throw new Exception(
                        "Bukan file reimbursement: sheet Resume/Rincian tidak ditemukan.");
            }

            for (String alamat : new String[]{"G10", "G11", "G12", "H20"}) {
                CellAddress address = new CellAddress(alamat);
                data.resume.put(alamat,
                        teksSel(resume, address.getRow(), address.getColumn()));
            }

            bacaBaris(rincian, 9, 37, data.material);
            bacaBaris(rincian, 43, 49, data.jasa);
            bacaBaris(rincian, 55, 117, data.operasional);

            Sheet lampiran = workbook.getSheet("Lampiran");

            if (lampiran instanceof XSSFSheet) {
                XSSFDrawing drawing = ((XSSFSheet) lampiran).getDrawingPatriarch();

                if (drawing != null) {
                    for (XSSFShape shape : drawing.getShapes()) {
                        if (!(shape instanceof XSSFPicture)) continue;

                        XSSFPicture gambar = (XSSFPicture) shape;
                        byte[] isi = gambar.getPictureData().getData();

                        if (isi == null || isi.length == 0) continue;

                        // Keterangan ditulis satu baris di atas foto.
                        String keterangan = "";
                        try {
                            XSSFClientAnchor anchor = gambar.getClientAnchor();
                            keterangan = teksSel(lampiran, anchor.getRow1() - 1, 0);
                        } catch (Throwable ignored) {
                            // File lama tanpa keterangan.
                        }

                        data.photos.add(new Gambar(isi, 0, 0, keterangan));
                    }
                }
            }
        }

        return data;
    }

    /** Membaca kolom A-E dari baris pertama s/d terakhir (nomor baris Excel). */
    private static void bacaBaris(
            Sheet sheet,
            int firstRow,
            int lastRow,
            List<List<String>> result
    ) {
        for (int r = firstRow - 1; r < lastRow; r++) {
            List<String> values = new ArrayList<>();

            for (int c = 0; c < 5; c++) {
                values.add(teksSel(sheet, r, c));
            }

            result.add(values);
        }
    }

    // =====================================================================
    // Menulis
    // =====================================================================

    public static void writeWorkbook(
            InputStream templateInput,
            OutputStream output,
            Map<String, Map<String, Object>> allValues,
            List<Gambar> photos
    ) throws Exception {

        try (Workbook workbook = WorkbookFactory.create(templateInput)) {

            // 1. Tulis data sel pada semua sheet.
            for (Map.Entry<String, Map<String, Object>> sheetEntry
                    : allValues.entrySet()) {

                Sheet sheet = workbook.getSheet(sheetEntry.getKey());

                if (sheet == null) {
                    throw new IllegalArgumentException(
                            "Sheet tidak ditemukan: " + sheetEntry.getKey());
                }

                for (Map.Entry<String, Object> entry
                        : sheetEntry.getValue().entrySet()) {

                    CellAddress address = new CellAddress(entry.getKey());
                    tulisSel(sheet, address.getRow(), address.getColumn(),
                            entry.getValue());
                }
            }

            // 2. Sisipkan foto bon pada sheet Lampiran, masing-masing
            //    dengan keterangan di atasnya dan proporsi aslinya.
            if (photos != null && !photos.isEmpty()) {
                Sheet lampiran = workbook.getSheet("Lampiran");

                if (lampiran == null) {
                    throw new IllegalArgumentException(
                            "Sheet Lampiran tidak ditemukan.");
                }

                Drawing<?> drawing = lampiran.createDrawingPatriarch();
                CreationHelper helper = workbook.getCreationHelper();

                float tinggiBaris = lampiran.getDefaultRowHeightInPoints() * 96f / 72f;
                if (tinggiBaris <= 0) tinggiBaris = 20f;

                int baris = 0;

                for (Gambar foto : photos) {
                    if (foto == null || foto.data == null || foto.data.length == 0) {
                        continue;
                    }

                    // Keterangan foto.
                    tulisSel(lampiran, baris, 0, foto.keterangan);

                    int pictureType = isPng(foto.data)
                            ? Workbook.PICTURE_TYPE_PNG
                            : Workbook.PICTURE_TYPE_JPEG;

                    int pictureIndex = workbook.addPicture(foto.data, pictureType);

                    // Hitung ukuran agar foto tidak gepeng.
                    float rasio = foto.lebar > 0 && foto.tinggi > 0
                            ? (float) foto.lebar / foto.tinggi
                            : 0.75f;

                    float tinggi = TINGGI_FOTO;
                    float lebar = tinggi * rasio;

                    if (lebar > LEBAR_MAKS_FOTO) {
                        lebar = LEBAR_MAKS_FOTO;
                        tinggi = lebar / rasio;
                    }

                    int barisAwal = baris + 1;

                    ClientAnchor anchor = helper.createClientAnchor();
                    anchor.setCol1(0);
                    anchor.setRow1(barisAwal);
                    anchor.setDx1(0);
                    anchor.setDy1(0);

                    // Kolom akhir + sisa piksel.
                    int kolom = 0;
                    float sisaLebar = lebar;
                    while (true) {
                        float lebarKolom = lampiran.getColumnWidthInPixels(kolom);
                        if (lebarKolom <= 0) lebarKolom = 64f;
                        if (sisaLebar <= lebarKolom || kolom > 50) break;
                        sisaLebar -= lebarKolom;
                        kolom++;
                    }
                    anchor.setCol2(kolom);
                    anchor.setDx2(Math.round(sisaLebar * EMU_PER_PIKSEL));

                    // Baris akhir + sisa piksel.
                    int jumlahBarisPenuh = (int) (tinggi / tinggiBaris);
                    float sisaTinggi = tinggi - jumlahBarisPenuh * tinggiBaris;
                    anchor.setRow2(barisAwal + jumlahBarisPenuh);
                    anchor.setDy2(Math.round(sisaTinggi * EMU_PER_PIKSEL));

                    anchor.setAnchorType(ClientAnchor.AnchorType.MOVE_AND_RESIZE);

                    drawing.createPicture(anchor, pictureIndex);

                    // Foto berikutnya: beri jarak 2 baris.
                    baris = barisAwal + jumlahBarisPenuh + 3;
                }
            }

            // 3. Hitung ulang rumus template (total, saldo) agar nilainya
            //    langsung tampil, dan minta Excel menghitung ulang saat dibuka.
            workbook.setForceFormulaRecalculation(true);

            try {
                workbook.getCreationHelper()
                        .createFormulaEvaluator()
                        .evaluateAll();
            } catch (Throwable t) {
                // Tidak fatal: Excel tetap menghitung ulang saat dibuka.
                android.util.Log.w("REMBUSE_EXCEL", "Evaluasi rumus dilewati", t);
            }

            // 4. Simpan workbook sekali saja.
            workbook.write(output);
            output.flush();
        }
    }

    private static void tulisSel(Sheet sheet, int baris, int kolom, Object value) {
        Row row = sheet.getRow(baris);
        if (row == null) row = sheet.createRow(baris);

        Cell cell = row.getCell(kolom);
        if (cell == null) cell = row.createCell(kolom);

        if (value == null) {
            cell.setBlank();
        } else if (value instanceof Number) {
            cell.setCellValue(((Number) value).doubleValue());
        } else if (value instanceof Boolean) {
            cell.setCellValue((Boolean) value);
        } else {
            cell.setCellValue(value.toString());
        }
    }

    private static boolean isPng(byte[] data) {
        return data.length >= 8
                && (data[0] & 0xFF) == 0x89
                && data[1] == 0x50
                && data[2] == 0x4E
                && data[3] == 0x47
                && data[4] == 0x0D
                && data[5] == 0x0A
                && data[6] == 0x1A
                && data[7] == 0x0A;
    }
}
