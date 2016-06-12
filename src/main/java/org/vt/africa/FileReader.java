package org.vt.africa;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DateFormatSymbols;
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
    private static final int MEAN_SUM_INDEX = LAST_DAY_MEASUREMENT + 1;
    private static final int ABS_DIFF_INDEX = LAST_DAY_MEASUREMENT + 2;
    private static final int PERCENT_DIFF_INDEX = LAST_DAY_MEASUREMENT + 3;
    private static final int MONTH_NAME_INDEX = LAST_DAY_MEASUREMENT + 4;
    private static final int PREC_RAW_VALUES_INDEX = LAST_DAY_MEASUREMENT + 5;
    private static final int MONTH_COLUMN = 8;
    private static final int YEAR_COLUMN = 7;
    private static final int MONTH_CELL_COLUMN = 8;
    private static final String SUM_COLUMN_NAME = "AP";
    private static final String MEAN_COLUMN_NAME = "AQ";
    private static final int MEAN_YEAR = 3;
    private static final int MAX_MONTH = 12;

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
                String sumFormula = "SUM(" + SUM_COLUMN_NAME + index + ":" + SUM_COLUMN_NAME + i + ")";
                sumColumn.setCellFormula(sumFormula);
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
            int month = 0;
            if (monthCell != null) {
                stringMonthValue = monthCell.getStringCellValue();
                month = Integer.valueOf(stringMonthValue);
            }
            Cell yearCell = row.getCell(YEAR_COLUMN);
            int year = 0;
            if (yearCell != null) {
                String stringYearValue = yearCell.getStringCellValue();
                year = Integer.valueOf(stringYearValue);
            }
            sumList.add(new MonthInYear(month, year, sumForMonth));
            System.out.println("month: " + month + " year: " + year + " SUM: " + sumForMonth);
        }

        List<Double> meanList = calculateMeanValues(sheet, firstRow, sumList);
        calculateAbsDiffValues(sheet, firstRow);
        calculatePercentDiffValues(sheet, firstRow);
        fillMonthNames(sheet, firstRow);
        fillPrecipitationRawValues(sheet, firstRow, meanList);
        createMeanSumFormula(sheet.getRow(MAX_MONTH + 1));
        createDiagram(workbook, sheet, meanList);

        writeOutput(file, workbook);
    }

    private void writeOutput(File file, XSSFWorkbook workbook) {
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

    private void createMeanSumFormula(Row row) {
        Cell meanCellSum = row.createCell(MEAN_SUM_INDEX);
        String beginMeanValue = String.valueOf(2);
        String endMeanValue = String.valueOf(13);
        String expectedFormula = "SUM(" + MEAN_COLUMN_NAME + beginMeanValue + ":" + MEAN_COLUMN_NAME + endMeanValue + ")";
        meanCellSum.setCellFormula(expectedFormula);
    }

    private List<Double> calculateMeanValues(XSSFSheet sheet, Row firstRow, List<MonthInYear> sumList) {
        Cell meanCellTitle = firstRow.createCell(MEAN_SUM_INDEX);
        meanCellTitle.setCellValue("Mean of " + MEAN_YEAR + " years");
        List<Double> meanList = new ArrayList<>();
        for (int i = 1; i <= MAX_MONTH; i++) {
            Row row = sheet.getRow(i);
            Cell meanCell = row.createCell(MEAN_SUM_INDEX);
            String firstYearIndex = String.valueOf(i + 1);
            String secondYearIndex = String.valueOf(i + 14);
            String meanFormula;
            double meanValue;
            if (MEAN_YEAR == 3) {
                String thirdYearIndex = String.valueOf(i + 27);
                meanFormula = "(" + SUM_COLUMN_NAME + firstYearIndex + "+" +
                        SUM_COLUMN_NAME + secondYearIndex + "+" +
                        SUM_COLUMN_NAME + thirdYearIndex + ")" +
                        "/" + MEAN_YEAR;
                double sumForFirstYear = sumList.get(i - 1).getSum();
                double sumForSecondYear = sumList.get(i + 11).getSum();
                double sumForThirdYear = sumList.get(i + 23).getSum();
                meanValue = (sumForFirstYear + sumForSecondYear + sumForThirdYear) / MEAN_YEAR;
            } else if (MEAN_YEAR == 2) {
                meanFormula = "(" + SUM_COLUMN_NAME + firstYearIndex + "+" +
                        SUM_COLUMN_NAME + secondYearIndex + ")" +
                        "/" + MEAN_YEAR;
                double sumForFirstYear = sumList.get(i - 1).getSum();
                double sumForSecondYear = sumList.get(i + 11).getSum();
                meanValue = (sumForFirstYear + sumForSecondYear) / MEAN_YEAR;
            }
            meanCell.setCellFormula(meanFormula);
            meanList.add(meanValue);
        }
        return meanList;
    }

    private void calculateAbsDiffValues(XSSFSheet sheet, Row firstRow) {
        Cell meanCellTitle = firstRow.createCell(ABS_DIFF_INDEX);
        meanCellTitle.setCellValue("Abs. diff.");
        for (int i = 1; i <= MAX_MONTH; i++) {
            Row row = sheet.getRow(i);
            Cell absDiffCell = row.createCell(ABS_DIFF_INDEX);
            String firstYearIndex = String.valueOf(i + 1);
            String secondYearIndex = String.valueOf(i + 14);
            String absDiffFormula;
            if (MEAN_YEAR == 3) {
                String thirdYearIndex = String.valueOf(i + 27);
                absDiffFormula = "MAX(" + SUM_COLUMN_NAME + firstYearIndex + "," +
                        SUM_COLUMN_NAME + secondYearIndex + "," +
                        SUM_COLUMN_NAME + thirdYearIndex + ")" +
                            "-" +
                        "MIN(" + SUM_COLUMN_NAME + firstYearIndex + "," +
                        SUM_COLUMN_NAME + secondYearIndex + "," +
                        SUM_COLUMN_NAME + thirdYearIndex + ")";
            } else if (MEAN_YEAR == 2) {
                absDiffFormula = SUM_COLUMN_NAME + firstYearIndex + "-" + SUM_COLUMN_NAME + secondYearIndex;
            }
            absDiffCell.setCellFormula(absDiffFormula);
        }
    }

    private void calculatePercentDiffValues(XSSFSheet sheet, Row firstRow) {
        Cell percentDiffTitle = firstRow.createCell(PERCENT_DIFF_INDEX);
        percentDiffTitle.setCellValue("Diff in %");
        for (int i = 1; i <= MAX_MONTH; i++) {
            Row row = sheet.getRow(i);
            Cell percentDiffCell = row.createCell(PERCENT_DIFF_INDEX);
            String sameMonthValue = String.valueOf(i + 1);
            String percentDiffFormula = "IFERROR((AR" + sameMonthValue + "*100)/AQ" + sameMonthValue + ", 0)";
            percentDiffCell.setCellFormula(percentDiffFormula);
        }
    }

    private void fillMonthNames(XSSFSheet sheet, Row firstRow) {
        Cell monthTitle = firstRow.createCell(MONTH_NAME_INDEX);
        monthTitle.setCellValue("Month");
        String[] months = new DateFormatSymbols().getMonths();
        for (int i = 1; i <= MAX_MONTH; i++) {
            Row row = sheet.getRow(i);
            Cell monthNameCell = row.createCell(MONTH_NAME_INDEX);
            monthNameCell.setCellValue(months[i - 1]);
        }
    }

    private void fillPrecipitationRawValues(XSSFSheet sheet, Row firstRow, List<Double> meanList) {
        Cell precRawIndex = firstRow.createCell(PREC_RAW_VALUES_INDEX);
        precRawIndex.setCellValue("Precip (mm)");
        for (int i = 1; i <= MAX_MONTH; i++) {
            Row row = sheet.getRow(i);
            Cell precValueCell = row.createCell(PREC_RAW_VALUES_INDEX);
            precValueCell.setCellValue(meanList.get(i - 1));
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
        for (int i = 1; i <= MAX_MONTH; i++) {
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
