import org.apache.commons.compress.archivers.zip.Zip64Mode;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

public class POIZipTest {
    public static void main(String[] args) {
            String path = "D:\\tmp\\test.xlsx";
            try {
                createExcel(path);  // create Excel document (which is a ZIP)
                showZipContents(path);  // try to open Excel document using java.util.zip -> exception
            } catch (Exception e) {
                e.printStackTrace();
            }
    }

    public static void showZipContents(String zipFileName) throws FileNotFoundException, IOException {
        FileInputStream inputStream = new FileInputStream(new File(zipFileName));
        ZipInputStream zip = new ZipInputStream(inputStream);
        ZipEntry entry = null;
        while ((entry = zip.getNextEntry()) != null) {
            System.out.println(entry.getName());
        }
        inputStream.close();
    }

    public static void createExcel(String fileName) throws IOException {
        SXSSFWorkbook workbook = new SXSSFWorkbook(); // streaming xssf workbook enables Zip64Mode.AsNeeded
        SXSSFSheet sheet = workbook.createSheet("Java Books");
        // workbook.setZip64Mode(Zip64Mode.Never);  // must disable Zip64 to read this with java.util.zip

        Object[][] bookData = {
                {"Head First Java", "Kathy Serria", 79},
                {"Effective Java", "Joshua Bloch", 36},
                {"Clean Code", "Robert martin", 42},
                {"Thinking in Java", "Bruce Eckel", 35},
        };

        int rowCount = 0;

        for (Object[] aBook : bookData) {
            Row row = sheet.createRow(++rowCount);

            int columnCount = 0;

            for (Object field : aBook) {
                Cell cell = row.createCell(++columnCount);
                if (field instanceof String) {
                    cell.setCellValue((String) field);
                } else if (field instanceof Integer) {
                    cell.setCellValue((Integer) field);
                }
            }
        }
        FileOutputStream outputStream = null;
        try {
            outputStream = new FileOutputStream(fileName);
            workbook.write(outputStream);
        } catch (Exception e) {
            e.printStackTrace();
        }
        outputStream.close();
    }
}