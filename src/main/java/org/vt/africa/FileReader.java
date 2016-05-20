package org.vt.africa;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jfree.chart.ChartFactory;
import org.jfree.chart.ChartUtilities;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.plot.PlotOrientation;
import org.jfree.data.category.DefaultCategoryDataset;

import java.io.*;

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
            int lastRowNum = sheet.getLastRowNum();
            Row firstRow = sheet.getRow(0);
            Cell sumCell = firstRow.createCell(LAST_DAY_MEASUREMENT);
            sumCell.setCellValue("Sum");
            for (int i = 1; i <= lastRowNum + 2; i++) {
                Row row = sheet.getRow(i);
                if (row == null) {
                    break;
                }
                if (i == 13 || i == 26) { // all +13, skip one line
                    sheet.shiftRows(i, lastRowNum + 1, 1);
                    Row newRow = sheet.createRow(i);
                    Cell sumColumn = newRow.createCell(LAST_DAY_MEASUREMENT);
                    String index = String.valueOf(i - 11);
                    String expectedFormula = "SUM(AP" + index + ":AP" + i + ")";
                    sumColumn.setCellFormula(expectedFormula);
                    continue;
                }
                double sumForMonth = 0.0;
                for (int j = FIRST_DAY_MEASUREMENT; j < LAST_DAY_MEASUREMENT; j++) {
                    Cell column = row.getCell(j);
                    double numericCellValue = column.getNumericCellValue();
                    sumForMonth += numericCellValue;
                }
                Cell cell = row.getCell(LAST_DAY_MEASUREMENT);
                if (cell == null || "".equals(cell.toString())) {
                    Cell sumColumn = row.createCell(LAST_DAY_MEASUREMENT);
                    String index = String.valueOf(i + 1);
                    String expectedFormula = "SUM(K" + index + ":AO" + index + ")";
                    sumColumn.setCellFormula(expectedFormula);
                }
                System.out.println("month: " + row.getCell(MONTH_COLUMN) + " year: " + row.getCell(YEAR_COLUMN) + " SUM: " + sumForMonth);
            }
            Cell meanCellTitle = firstRow.createCell(LAST_DAY_MEASUREMENT + 1);
            meanCellTitle.setCellValue("Mean of 2 years");
            for (int i = 1; i < 13; i++) {
                Row row = sheet.getRow(i);
                Cell meanCell = row.createCell(LAST_DAY_MEASUREMENT + 1);
                String index = String.valueOf(i + 1);
                String nextIndex = String.valueOf(i + 14);
                String expectedFormula = "(AP" + index + "+AP" + nextIndex + ")/2";
                meanCell.setCellFormula(expectedFormula);
            }
            Row row = sheet.getRow(13);
            Cell meanCellSum = row.createCell(LAST_DAY_MEASUREMENT + 1);
            String index = String.valueOf(2);
            String nextIndex = String.valueOf(13);
            String expectedFormula = "SUM(AQ" + index + ":AQ" + nextIndex + ")";
            meanCellSum.setCellFormula(expectedFormula);

            createDiagram(workbook, sheet);
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

    private void createDiagram(XSSFWorkbook workbook, XSSFSheet sheet) {
        DefaultCategoryDataset dataset = new DefaultCategoryDataset();
        for (int i = 1; i < 13; i++) {
            Row currentRow = sheet.getRow(i);
            Cell meanCell = currentRow.getCell(LAST_DAY_MEASUREMENT + 1);
            double value = meanCell.getNumericCellValue();
            dataset.addValue(value, "Marks", "1");
        }
        JFreeChart BarChartObject = ChartFactory.createBarChart(
                "Subject Vs Marks", "Subject", "Marks", dataset,
                PlotOrientation.VERTICAL, true, true, false);
        int width = 640;
        int height = 480;
        ByteArrayOutputStream chart_out = new ByteArrayOutputStream();
        try {
            ChartUtilities.writeChartAsPNG(chart_out,BarChartObject,width,height);
        } catch (IOException e) {
            e.printStackTrace();
        }
        int myPictureId = workbook.addPicture(chart_out.toByteArray(), Workbook.PICTURE_TYPE_PNG);
        try {
            chart_out.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        XSSFDrawing drawing = sheet.createDrawingPatriarch();
        ClientAnchor my_anchor = new XSSFClientAnchor();
        my_anchor.setCol1(4);
        my_anchor.setRow1(5);
        drawing.createPicture(my_anchor, myPictureId);
    }
}
