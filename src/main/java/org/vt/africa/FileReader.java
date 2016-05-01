package org.vt.africa;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * @author vladimir.tikhomirov
 */
public class FileReader {

    private static final int FIRST_DAY_MEASUREMENT = 10;
    private static final int LAST_DAY_MEASUREMENT = FIRST_DAY_MEASUREMENT + 31;
    private static final int MONTH_COLUMN = 8;
    private static final int YEAR_COLUMN = 7;

    public void fetchDataFromFile(File file) {
        FileInputStream inputStream = null;
        try {
            inputStream = new FileInputStream(file);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }

        XSSFWorkbook workbook = null;
        XSSFSheet sheet = null;
        if (inputStream != null) {
            try {
                workbook = new XSSFWorkbook(inputStream);
            } catch (IOException e) {
                e.printStackTrace();
            }
            if (workbook != null) {
                sheet = workbook.getSheetAt(0);
            }
            try {
                inputStream.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

        if (sheet != null) {
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                double sumForMonth = 0.0;
                for (int j = FIRST_DAY_MEASUREMENT; j < LAST_DAY_MEASUREMENT; j++) {
                    Cell column = row.getCell(j);
                    double numericCellValue = column.getNumericCellValue();
                    sumForMonth += numericCellValue;
                }
                Cell cell = row.getCell(LAST_DAY_MEASUREMENT);
                if (cell == null || "".equals(cell.toString())) {
                    Cell sumColumn = row.createCell(LAST_DAY_MEASUREMENT);
                    sumColumn.setCellValue(sumForMonth);
                }
                System.out.println("month: " + row.getCell(MONTH_COLUMN) + " year: " + row.getCell(YEAR_COLUMN) + " SUM: " + sumForMonth);
            }
        }

        FileOutputStream outFile = null;
        try {
            outFile = new FileOutputStream(new File(file.getPath()));
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        try {
            if (outFile != null) {
                if (workbook != null) {
                    workbook.write(outFile);
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                if (outFile != null) {
                    outFile.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }
}
