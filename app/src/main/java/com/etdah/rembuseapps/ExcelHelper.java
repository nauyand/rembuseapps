package com.etdah.rembuseapps;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.File;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.util.LinkedHashMap;
import java.util.Locale;
import java.util.List;
import java.util.Map;
import java.util.ArrayList;
public class ExcelHelper {

    /**
     * Mengambil teks sel. Sel bertipe tanggal selalu dikembalikan
     * sebagai dd/MM/yyyy agar tahunnya tidak hilang (misalnya jika
     * format sel di Excel hanya "d-mmm").
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

    /**
     * Angka ditulis apa adanya, tanpa format Excel.
     * Sengaja TIDAK memakai DataFormatter: pada format tertentu
     * (misalnya format akuntansi) DataFormatter memanggil java.awt.Color
     * yang tidak tersedia di Android sehingga aplikasi force close.
     */
    private static String angkaPolos(double angka) {
        if (angka == Math.rint(angka) && Math.abs(angka) < 1e15) {
            return String.valueOf((long) angka);
        }
        return java.math.BigDecimal.valueOf(angka)
                .stripTrailingZeros()
                .toPlainString();
    }

    /** Membaca beberapa sel dari satu sheet dengan sekali buka file. */
    public static Map<String, String> readCells(
            InputStream input,
            String sheetName,
            String... cellAddresses
    ) throws Exception {
        Map<String, String> result = new LinkedHashMap<>();

        try (Workbook workbook = WorkbookFactory.create(input)) {
            Sheet sheet = workbook.getSheet(sheetName);

            if (sheet == null) {
                throw new IllegalArgumentException(
                        "Sheet tidak ditemukan: " + sheetName);
            }


            for (String alamat : cellAddresses) {
                CellAddress address = new CellAddress(alamat);
                Row row = sheet.getRow(address.getRow());
                Cell cell = row == null
                        ? null
                        : row.getCell(address.getColumn());

                result.put(alamat, cellText(cell));
            }
        }

        return result;
    }

    /** Semua data yang dibaca dari file reimbursement. */
    public static class ReimburseData {
        public final Map<String, String> resume = new LinkedHashMap<>();
        public final List<List<String>> material = new ArrayList<>();
        public final List<List<String>> jasa = new ArrayList<>();
        public final List<List<String>> operasional = new ArrayList<>();
        public final List<byte[]> photos = new ArrayList<>();
    }

    /**
     * Membaca seluruh isi file reimbursement dengan SEKALI buka.
     * File dibuka langsung dari disk dalam mode baca-saja agar hemat
     * memori (penting jika ada banyak foto bon berukuran besar).
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
                Row row = resume.getRow(address.getRow());
                Cell cell = row == null ? null : row.getCell(address.getColumn());
                data.resume.put(alamat, cellText(cell));
            }

            bacaBaris(rincian, 9, 37, data.material);
            bacaBaris(rincian, 43, 49, data.jasa);
            bacaBaris(rincian, 55, 117, data.operasional);

            Sheet lampiran = workbook.getSheet("Lampiran");

            if (lampiran instanceof XSSFSheet) {
                XSSFDrawing drawing = ((XSSFSheet) lampiran).getDrawingPatriarch();

                if (drawing != null) {
                    for (XSSFShape shape : drawing.getShapes()) {
                        if (shape instanceof XSSFPicture) {
                            byte[] imageData = ((XSSFPicture) shape)
                                    .getPictureData()
                                    .getData();

                            if (imageData != null && imageData.length > 0) {
                                data.photos.add(imageData);
                            }
                        }
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
            Row row = sheet.getRow(r);

            for (int c = 0; c < 5; c++) {
                Cell cell = row == null ? null : row.getCell(c);
                values.add(cellText(cell));
            }

            result.add(values);
        }
    }

    public static List<byte[]> readPictures(
            InputStream input,
            String sheetName
    ) throws Exception {

        List<byte[]> photos = new ArrayList<>();

        try (Workbook workbook = WorkbookFactory.create(input)) {
            if (!(workbook instanceof XSSFWorkbook)) {
                throw new Exception("File harus berformat Excel .xlsx");
            }

            XSSFSheet sheet =
                    ((XSSFWorkbook) workbook).getSheet(sheetName);

            if (sheet == null) {
                throw new Exception(
                        "Sheet tidak ditemukan: " + sheetName
                );
            }

            XSSFDrawing drawing = sheet.getDrawingPatriarch();

            if (drawing == null) {
                return photos;
            }

            for (XSSFShape shape : drawing.getShapes()) {
                if (shape instanceof XSSFPicture) {
                    byte[] imageData =
                            ((XSSFPicture) shape)
                                    .getPictureData()
                                    .getData();

                    if (imageData != null && imageData.length > 0) {
                        photos.add(imageData);
                    }
                }
            }
        }

        return photos;
    }

    public static List<List<String>> readRange(
            InputStream input,
            String sheetName,
            int firstRow,
            int lastRow,
            int firstColumn,
            int lastColumn
    ) throws Exception {

        List<List<String>> result = new ArrayList<>();

        try (Workbook workbook = WorkbookFactory.create(input)) {
            Sheet sheet = workbook.getSheet(sheetName);

            if (sheet == null) {
                throw new IllegalArgumentException(
                        "Sheet tidak ditemukan: " + sheetName
                );
            }


            // Nomor baris dan kolom menggunakan indeks Excel mulai dari 1.
            for (int r = firstRow - 1; r < lastRow; r++) {
                List<String> values = new ArrayList<>();
                Row row = sheet.getRow(r);

                for (int c = firstColumn - 1; c < lastColumn; c++) {
                    Cell cell = row == null ? null : row.getCell(c);

                    values.add(cellText(cell));
                }

                result.add(values);
            }
        }

        return result;
    }

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
            Row row = sheet.getRow(address.getRow());

            if (row == null) return "";

            Cell cell = row.getCell(address.getColumn());

            return cellText(cell);
        }
    }

    public static void writeCells(
            InputStream templateInput,
            OutputStream output,
            String sheetName,
            Map<String, Object> values
    ) throws Exception {
        Map<String, Map<String, Object>> allValues =
                new LinkedHashMap<>();

        allValues.put(sheetName, values);

        writeWorkbook(templateInput, output, allValues, null);
    }

    public static void writeMultipleSheets(
            InputStream templateInput,
            OutputStream output,
            Map<String, Map<String, Object>> allValues
    ) throws Exception {
        writeWorkbook(templateInput, output, allValues, null);
    }

    public static void writeWorkbook(
            InputStream templateInput,
            OutputStream output,
            Map<String, Map<String, Object>> allValues,
            List<byte[]> photos
    ) throws Exception {

        try (Workbook workbook =
                     WorkbookFactory.create(templateInput)) {

            // 1. Tulis data sel pada semua sheet.
            for (Map.Entry<String, Map<String, Object>> sheetEntry
                    : allValues.entrySet()) {

                Sheet sheet = workbook.getSheet(sheetEntry.getKey());

                if (sheet == null) {
                    throw new IllegalArgumentException(
                            "Sheet tidak ditemukan: "
                                    + sheetEntry.getKey());
                }

                for (Map.Entry<String, Object> entry
                        : sheetEntry.getValue().entrySet()) {

                    CellAddress address =
                            new CellAddress(entry.getKey());

                    Row row = sheet.getRow(address.getRow());

                    if (row == null) {
                        row = sheet.createRow(address.getRow());
                    }

                    Cell cell = row.getCell(address.getColumn());

                    if (cell == null) {
                        cell = row.createCell(address.getColumn());
                    }

                    Object value = entry.getValue();

                    if (value == null) {
                        cell.setBlank();
                    } else if (value instanceof Number) {
                        cell.setCellValue(
                                ((Number) value).doubleValue());
                    } else if (value instanceof Boolean) {
                        cell.setCellValue((Boolean) value);
                    } else {
                        cell.setCellValue(value.toString());
                    }
                }
            }

            // 2. Sisipkan foto bon pada sheet Lampiran.
            if (photos != null && !photos.isEmpty()) {
                Sheet lampiran = workbook.getSheet("Lampiran");

                if (lampiran == null) {
                    throw new IllegalArgumentException(
                            "Sheet Lampiran tidak ditemukan.");
                }

                Drawing<?> drawing = lampiran.createDrawingPatriarch();
                CreationHelper helper = workbook.getCreationHelper();

                for (int i = 0; i < photos.size(); i++) {
                    byte[] imageBytes = photos.get(i);

                    if (imageBytes == null || imageBytes.length == 0) {
                        continue;
                    }

                    int pictureType = isPng(imageBytes)
                            ? Workbook.PICTURE_TYPE_PNG
                            : Workbook.PICTURE_TYPE_JPEG;

                    int pictureIndex = workbook.addPicture(
                            imageBytes, pictureType);

                    ClientAnchor anchor =
                            helper.createClientAnchor();

                    // Setiap foto mendapat area sendiri:
                    // kolom A-F, dengan jarak 20 baris.
                    int startRow = i * 20;

                    anchor.setCol1(0);
                    anchor.setRow1(startRow);
                    anchor.setCol2(6);
                    anchor.setRow2(startRow + 16);
                    anchor.setAnchorType(
                            ClientAnchor.AnchorType.MOVE_AND_RESIZE);

                    drawing.createPicture(anchor, pictureIndex);
                }
            }

            // 3. Hitung ulang rumus template (total, saldo) agar
            //    nilainya langsung tampil, dan minta Excel menghitung
            //    ulang saat file dibuka.
            workbook.setForceFormulaRecalculation(true);

            try {
                workbook.getCreationHelper()
                        .createFormulaEvaluator()
                        .evaluateAll();
            } catch (Throwable t) {
                // Tidak fatal: Excel tetap menghitung ulang saat dibuka.
                android.util.Log.w(
                        "REMBUSE_EXCEL",
                        "Evaluasi rumus dilewati",
                        t);
            }

            // 4. Simpan workbook sekali saja.
            workbook.write(output);
            output.flush();
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