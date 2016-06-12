package org.vt.africa;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jfree.chart.ChartFactory;
import org.jfree.chart.ChartUtilities;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.plot.PlotOrientation;
import org.jfree.data.category.DefaultCategoryDataset;

/**
 * @author vladimir.tikhomirov
 */
public class FileReader {

    private static final int FIRST_DAY_MEASUREMENT = 10;
    private static final int LAST_DAY_MEASUREMENT = FIRST_DAY_MEASUREMENT + 31;
    private static final int MONTH_COLUMN = 8;
    private static final int YEAR_COLUMN = 7;
    private static final int MONTH_CELL_COLUMN = 8;
    private static final int MEAN_YEAR = 3;

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

        if (sheet == null) {
            return;
        }


        int rowAmountToSkip = colorMissingMonthRows(workbook.createCellStyle(), sheet);
        List<Integer> correctRowNumbers = new ArrayList<>();
        int sum = 0;
        for (int i = 0; i < rowAmountToSkip; i++) {
            sum += 13;
            correctRowNumbers.add(sum);
        }

        XSSFCellStyle yellowStyle = workbook.createCellStyle();

        int lastRowNum = sheet.getLastRowNum();
        Row firstRow = sheet.getRow(0);
        Cell sumCell = firstRow.createCell(LAST_DAY_MEASUREMENT);
        sumCell.setCellValue("Sum");
        List<MonthInYear> sumList = new ArrayList<>();
        for (int i = 1; i <= lastRowNum; i++) {
            Row row = sheet.getRow(i);
            if (row == null) {
                break;
            }
            if (correctRowNumbers.contains(i)) {
                lastRowNum++;
                sheet.shiftRows(i, lastRowNum, 1);
                Row newRow = sheet.createRow(i);
                yellowStyle.setFillBackgroundColor(IndexedColors.YELLOW.getIndex());
                yellowStyle.setFillPattern(CellStyle.BIG_SPOTS);
                newRow.setRowStyle(yellowStyle);
                Cell sumColumn = newRow.createCell(LAST_DAY_MEASUREMENT);
                String index = String.valueOf(i - 11);
                String expectedFormula = "SUM(AP" + index + ":AP" + i + ")";
                sumColumn.setCellFormula(expectedFormula);
                continue;
            }
            double sumForMonth = 0.0;
            for (int j = FIRST_DAY_MEASUREMENT; j < LAST_DAY_MEASUREMENT; j++) {
                Cell column = row.getCell(j);
                if (column != null) {
                    double numericCellValue = column.getNumericCellValue();
                    sumForMonth += numericCellValue;
                }
            }
            Cell cell = row.getCell(LAST_DAY_MEASUREMENT);
            if (cell == null || "".equals(cell.toString())) {
                Cell sumColumn = row.createCell(LAST_DAY_MEASUREMENT);
                String index = String.valueOf(i + 1);
                String expectedFormula = "SUM(K" + index + ":AO" + index + ")";
                sumColumn.setCellFormula(expectedFormula);
            }
            Cell monthCell = row.getCell(MONTH_COLUMN);
            String stringMonthValue;
            double month = 0.0;
            if (monthCell != null) {
                stringMonthValue = monthCell.getStringCellValue();
                month = Double.valueOf(stringMonthValue);
            }
            Cell yearCell = row.getCell(YEAR_COLUMN);
            double year = 0.0;
            if (yearCell != null) {
                String stringYearValue = yearCell.getStringCellValue();
                year = Double.valueOf(stringYearValue);
            }
            sumList.add(new MonthInYear(month, year, sumForMonth));
            System.out.println("month: " + month + " year: " + year + " SUM: " + sumForMonth);
        }
        Cell meanCellTitle = firstRow.createCell(LAST_DAY_MEASUREMENT + 1);
        meanCellTitle.setCellValue("Mean of " + MEAN_YEAR + " years");
        List<Double> meanList = new ArrayList<>();
        for (int i = 1; i < 13; i++) {
            Row row = sheet.getRow(i);
            Cell meanCell = row.createCell(LAST_DAY_MEASUREMENT + 1);
            String index = String.valueOf(i + 1);
            String nextIndex = String.valueOf(i + 14);
            String expectedFormula;
            double expectedResult;
            if (MEAN_YEAR == 3) {
                String lastIndex = String.valueOf(i + 27);
                expectedFormula = "(AP" + index + "+AP" + nextIndex + "+AP" + lastIndex + ")/" + MEAN_YEAR;
                double sumForFirstYear = sumList.get(i - 1).getSum();
                double sumForSecondYear = sumList.get(i + 11).getSum();
                double sumForThirdYear = sumList.get(i + 23).getSum();
                expectedResult = (sumForFirstYear + sumForSecondYear + sumForThirdYear) / MEAN_YEAR;
            } else if (MEAN_YEAR == 2) {
                expectedFormula = "(AP" + index + "+AP" + nextIndex + ")/" + MEAN_YEAR;
                double sumForFirstYear = sumList.get(i - 1).getSum();
                double sumForSecondYear = sumList.get(i + 11).getSum();
                expectedResult = (sumForFirstYear + sumForSecondYear) / MEAN_YEAR;
            }
            meanCell.setCellFormula(expectedFormula);
            meanList.add(expectedResult);
        }
        Row row = sheet.getRow(13);
        Cell meanCellSum = row.createCell(LAST_DAY_MEASUREMENT + 1);
        String index = String.valueOf(2);
        String nextIndex = String.valueOf(13);
        String expectedFormula = "SUM(AQ" + index + ":AQ" + nextIndex + ")";
        meanCellSum.setCellFormula(expectedFormula);

        createDiagram(workbook, sheet, meanList);

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

    private int colorMissingMonthRows(CellStyle style, XSSFSheet sheet) {
        int lastRowNum = sheet.getLastRowNum();
        double expectedMonthValue = 1;
        int rowAmount = 0;
        for (int i = 1; i <= lastRowNum; i++) {
            Row row = sheet.getRow(i);
            if (row == null) {
                break;
            }
            Cell month = row.getCell(MONTH_CELL_COLUMN);
            if (month == null) {
                break;
            }
            String actualMonthValue = month.getStringCellValue();
            boolean fullYear = true;
            if (expectedMonthValue != Double.valueOf(actualMonthValue)) {
                fullYear = false;
                lastRowNum++;
                sheet.shiftRows(i, lastRowNum, 1);
                Row newRow = sheet.createRow(i);
                style.setFillBackgroundColor(IndexedColors.RED.getIndex());
                style.setFillPattern(CellStyle.BIG_SPOTS);
                newRow.setRowStyle(style);
            }
            expectedMonthValue++;
            if (expectedMonthValue == 13) {
                expectedMonthValue = 1;
                if (fullYear) {
                    rowAmount++;
                }
            }
        }
        return rowAmount;
    }

    private void createDiagram(XSSFWorkbook workbook, XSSFSheet sheet, List<Double> meanList) {
        DefaultCategoryDataset data = new DefaultCategoryDataset();
        for (int i = 1; i < 13; i++) {
            data.addValue(meanList.get(i - 1), "Mean value", i + "");
        }

        JFreeChart BarChartObject = ChartFactory.createBarChart("Mean of " + MEAN_YEAR + " years", "Month", "Mean value",
                data, PlotOrientation.VERTICAL, true, true, false);
        int width = 640;
        int height = 480;

        ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
        try {
            ChartUtilities.writeChartAsPNG(byteArrayOutputStream, BarChartObject, width, height);
        } catch (IOException e) {
            e.printStackTrace();
        }
        int pictureId = workbook.addPicture(byteArrayOutputStream.toByteArray(), Workbook.PICTURE_TYPE_PNG);
        try {
            byteArrayOutputStream.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        XSSFDrawing drawing = sheet.createDrawingPatriarch();

        ClientAnchor anchor = new XSSFClientAnchor();
        anchor.setCol1(43);
        anchor.setRow1(16);
        XSSFPicture picture = drawing.createPicture(anchor, pictureId);
        picture.resize();
    }
}
