package com.etdah.rembuseapps;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.ArrayList;
public class ExcelHelper {

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
                    ((XSSFWorkbook) workbook).getSheet("Lampiran");

            if (sheet == null) {
                throw new Exception(
                        "Sheet tidak ditemukan: " + "Lampiran"
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

            DataFormatter formatter = new DataFormatter();

            // Nomor baris dan kolom menggunakan indeks Excel mulai dari 1.
            for (int r = firstRow - 1; r < lastRow; r++) {
                List<String> values = new ArrayList<>();
                Row row = sheet.getRow(r);

                for (int c = firstColumn - 1; c < lastColumn; c++) {
                    Cell cell = row == null ? null : row.getCell(c);

                    values.add(
                            cell == null
                                    ? ""
                                    : formatter.formatCellValue(cell).trim()
                    );
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

            if (cell == null) return "";

            return new DataFormatter().formatCellValue(cell);
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

            // 3. Simpan workbook sekali saja.
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
