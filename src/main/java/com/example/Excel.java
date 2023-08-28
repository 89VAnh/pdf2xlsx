package com.example;

import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.TreeMap;

import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

public class Excel {
    public static CellStyle _defaultStyle;
    public static ArrayList<CellContent> countList = new ArrayList<>();
    public static HashMap<Integer, String> colColor = new HashMap<>();
    public static HashMap<Integer, String> rowColor = new HashMap<>();
    public static CellContent emptyCell = new CellContent();

    private static Cell getCell(Row row, int colIndex) {
        Cell cell;
        if (row.getCell(colIndex) == null) {
            cell = row.createCell(colIndex);
        } else {
            cell = row.getCell(colIndex);
        }
        return cell;
    }

    private static void setCell(Workbook workbook, Row row, int colIndex, CellContent c, String color) {
        if (colIndex >= 1 && row.getCell(colIndex) != null) {
            Cell cell = getCell(row, colIndex);
            String prevValue = cell.getStringCellValue();

            CellContent prevCell = new CellContent();
            prevCell.value = prevValue;
            prevCell.line_height = c.line_height;
            prevCell.font_family = c.font_family;
            prevCell.font_size = c.font_size;
            prevCell.color = c.color;

            setCell(workbook, row, colIndex - 1, prevCell, color);
            cell.setCellValue(c.value);
        } else {
            Cell cell = getCell(row, colIndex);
            cell.setCellValue(c.value);
            CellStyle cs = workbook.createCellStyle();

            cs.setBorderBottom(BorderStyle.THIN);
            cs.setBorderTop(BorderStyle.THIN);
            cs.setBorderRight(BorderStyle.THIN);
            cs.setBorderLeft(BorderStyle.THIN);

            int rowIndex = row.getRowNum();

            if (colIndex >= 4) {

                if (rowIndex < 9) {
                    cs.setVerticalAlignment(VerticalAlignment.CENTER);
                    cs.setAlignment(HorizontalAlignment.CENTER);
                } else {
                    cs.setVerticalAlignment(VerticalAlignment.CENTER);
                    cs.setAlignment(HorizontalAlignment.RIGHT);
                }
            }

            if (!color.isEmpty()) {
                if (colColor.get(colIndex) != null && colColor.get(colIndex).equals("#ffff99")) {
                    cs.setFillForegroundColor(hex2Index(colColor.get(colIndex)));
                } else {
                    cs.setFillForegroundColor(hex2Index(color));
                }

                cs.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            } else {

                if (colColor.get(colIndex) != null) {
                    if (rowColor.get(rowIndex) != null && !colColor.get(colIndex).equals("#ffff99")) {
                        cs.setFillForegroundColor(hex2Index(rowColor.get(rowIndex)));
                    } else {
                        cs.setFillForegroundColor(hex2Index(colColor.get(colIndex)));
                    }
                    cs.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                } else {
                    if (rowColor.get(rowIndex) != null) {
                        cs.setFillForegroundColor(hex2Index(rowColor.get(rowIndex)));
                        cs.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                    }
                }
            }

            if (c.font_family != null) {
                Font font = fontStyle(workbook, c.font_family, c.font_size, c.color);
                cs.setFont(font);
            }

            cell.setCellStyle(cs);
        }
    }

    private static Row getRow(Sheet sheet, int rowIndex) {
        Row row;
        if (sheet.getRow(rowIndex) == null) {
            row = sheet.createRow(rowIndex);
        } else {
            row = sheet.getRow(rowIndex);
        }
        return row;
    }

    private static boolean WriteRow(Workbook workbook, ArrayList<CellContent> contents, int startColumn,
            int rowIndex,
            HashMap<Double, Integer> leftColMap, CellStyle style) {
        Sheet sheet = workbook.getSheetAt(0);
        Row row = getRow(sheet, rowIndex);
        boolean flag = false;

        int l = leftColMap.size();
        double[] lefts = new double[l];
        int i = 0;
        for (double left : leftColMap.keySet()) {
            lefts[i] = left;
            i++;
        }
        Arrays.sort(lefts);
        double maxLeft = lefts[l - 1];

        for (CellContent cellcontent : contents) {
            double colIndex = (double) Math.round(cellcontent.left / 30);
            if (leftColMap.get(colIndex) != null) {
                int col = startColumn + leftColMap.get(colIndex);
                if (cellcontent.value.toLowerCase().equals("total")) {
                    int prevTotalCount = 1;

                    int prevRow = rowIndex - 1;
                    while (sheet.getRow(prevRow).getCell(col).getStringCellValue().isEmpty()) {
                        prevTotalCount++;
                        prevRow--;
                    }
                    int rowPlus = 0;
                    String color = "";
                    switch (prevTotalCount) {
                        case 1:
                            color = "#ccffff";
                            rowColor.put(rowIndex, color);
                            break;
                        case 2:
                            color = "#32CD32";
                            rowColor.put(rowIndex, color);
                            break;
                        case 3:
                            color = "#e2efda";
                            rowPlus = 1;
                            rowColor.put(rowIndex, color);
                            rowColor.put(rowIndex + rowPlus, color);
                            break;
                        default:
                            break;
                    }

                    WriteMergeCell(workbook, rowIndex, rowIndex + rowPlus, col - prevTotalCount, col, cellcontent,
                            color);
                } else {
                    setCell(workbook, row, col, cellcontent, "");
                }
                flag = true;
            } else {
                if (colIndex < maxLeft) {
                    countList.add(cellcontent);
                } else {
                    int col = leftColMap.get(maxLeft);
                    setCell(workbook, row, startColumn + col, cellcontent, "");
                    flag = true;
                }
            }
        }
        return flag;
    }

    private static void WriteMergeCell(Workbook workbook, int firstRow, int lastRow, int firstCol, int lastCol,
            CellContent cc, String backgroundColor) {

        Sheet sheet = workbook.getSheetAt(0);

        Row row = getRow(sheet, firstRow);
        Cell cell = getCell(row, firstCol);
        CellStyle cs = workbook.createCellStyle();
        cs.setAlignment(HorizontalAlignment.CENTER);
        cs.setVerticalAlignment(VerticalAlignment.CENTER);
        cs.setBorderBottom(BorderStyle.THIN);
        cs.setBorderTop(BorderStyle.THIN);
        cs.setBorderRight(BorderStyle.THIN);
        cs.setBorderLeft(BorderStyle.THIN);

        if (!backgroundColor.isEmpty()) {
            cs.setFillForegroundColor(hex2Index(backgroundColor));
            cs.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        }

        if (cc.isRotate)
            cs.setRotation((short) 90);

        if (cc.value != null) {
            cell.setCellValue(cc.value);
        }

        if (cc.font_family != null) {
            Font font = fontStyle(workbook, cc.font_family, cc.font_size, cc.color);
            cs.setFont(font);
        }

        cell.setCellStyle(cs);

        for (int r = firstRow; r <= lastRow; r++) {
            Row tmpRow = getRow(sheet, r);
            for (int c = firstCol; c <= lastCol; c++) {
                if (r == firstRow && c == firstCol) {
                    continue;
                }
                Cell tmpCell = getCell(tmpRow, c);
                CellStyle tmpCs = workbook.createCellStyle();
                tmpCs.setBorderBottom(BorderStyle.THIN);
                tmpCs.setBorderTop(BorderStyle.THIN);
                tmpCs.setBorderRight(BorderStyle.THIN);
                tmpCs.setBorderLeft(BorderStyle.THIN);
                tmpCell.setCellStyle(tmpCs);
            }
        }

        CellRangeAddress cellRange = new CellRangeAddress(firstRow, lastRow,
                firstCol, lastCol);
        sheet.addMergedRegion(cellRange);

    }

    private static HashMap<Double, Integer> GetLeftColMap(Sheet sheet, TreeMap<Integer, ArrayList<CellContent>> rows) {
        int firstTotalRow = 0;
        for (int j = 4; j < rows.size(); j++) {
            if (rows.get(j).get(0).value.toLowerCase().equals("total")) {
                firstTotalRow = j;
                break;
            }
        }

        HashMap<Double, Integer> leftColMap = new HashMap<Double, Integer>();
        ArrayList<CellContent> totalRow = rows.get(firstTotalRow);

        for (int i = 0; i < totalRow.size(); i++) {
            CellContent cell = totalRow.get(i);
            leftColMap.put((double) Math.round(cell.left / 30), i);
        }

        return leftColMap;
    }

    private static CellStyle defaultCellStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();

        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        return style;
    }

    public static short hex2Index(String color) {
        short palIndex = 1;
        try {
            HSSFWorkbook hssfWorkbook = new HSSFWorkbook();
            HSSFPalette palette = hssfWorkbook.getCustomPalette();
            color = color.replace("#", "");
            int resultRed = Integer.valueOf(color.substring(0, 2), 16);
            int resultGreen = Integer.valueOf(color.substring(2, 4), 16);
            int resultBlue = Integer.valueOf(color.substring(4, 6), 16);

            HSSFColor myColor = palette.findSimilarColor(resultRed, resultGreen,
                    resultBlue);
            palIndex = myColor.getIndex();
            hssfWorkbook.close();
        } catch (Exception e) {
            System.out.println(e.getMessage());
        }
        return palIndex;
    }

    private static Font fontStyle(Workbook workbook, String font_family, double size, String color) {
        Font font = workbook.createFont();

        font.setFontName(font_family);
        font.setFontHeightInPoints((short) (size * 3));
        font.setBold(font_family.contains("F1"));

        if (!color.isEmpty()) {
            font.setColor(hex2Index(color));
        }
        return font;
    }

    public static void Write(String fileName, TreeMap<Integer, ArrayList<CellContent>> treeMap,
            ArrayList<CellContent> rotateList, ArrayList<CellContent> facList) {

        int l = treeMap.size();
        int i = l;

        TreeMap<Integer, ArrayList<CellContent>> rows = new TreeMap<Integer, ArrayList<CellContent>>();

        for (ArrayList<CellContent> row : treeMap.descendingMap().values()) {
            rows.put(i, row);
            i--;
        }

        try {
            Workbook workbook = new XSSFWorkbook();
            Sheet sheet = workbook.createSheet();

            _defaultStyle = defaultCellStyle(workbook);

            HashMap<Double, Integer> leftColMap = GetLeftColMap(sheet, rows);

            colColor.put(4, "#bcd6ed");
            colColor.put(5, "#bcd6ed");
            colColor.put(6, "#bcd6ed");

            // row 1
            WriteRow(workbook, rows.get(1), 3, 0, leftColMap, _defaultStyle);

            // row 2
            int row1Size = rows.get(1).size();
            int row2Size = rows.get(2).size();
            int colNumContent = row1Size / row2Size + 1;
            int firstCol = 4;

            for (CellContent cellContent : rows.get(2)) {
                colColor.put(firstCol + colNumContent - 1, "#ffff99");
                WriteMergeCell(workbook, 1, 1, firstCol, firstCol + colNumContent - 1, cellContent, "");
                firstCol += colNumContent;
            }

            // row 3
            for (CellContent cellContent : rows.get(3)) {
                WriteMergeCell(workbook, 1, 2, firstCol, firstCol, cellContent, "");
                firstCol++;
            }

            int indexRow = 2;
            for (int j = 4; j < rows.size(); j++) {
                if (WriteRow(workbook, rows.get(j), 3, indexRow, leftColMap, _defaultStyle)) {
                    indexRow++;
                }
            }
            WriteRow(workbook, rows.get(treeMap.size()), 3, indexRow, leftColMap, _defaultStyle);

            WriteMergeCell(workbook, 1, 1, 0, 3, new CellContent(), "#ffcc99");
            WriteMergeCell(workbook, 2, 8, 0, 2, facList.get(0), "#ffcc99");

            ArrayList<CellContent> rotateArray = new ArrayList<CellContent>();
            for (CellContent cellContent : rotateList) {
                if (cellContent.value.startsWith("LINE"))
                    WriteMergeCell(workbook, 9, indexRow - 2, 0, 0, cellContent, "#e2efda");
                else {
                    rotateArray.add(cellContent);
                }
            }

            // col 1
            int rowCount = 0;
            int index = 0;
            for (int rowI = 9; rowI < indexRow - 1; rowI++) {
                if (sheet.getRow(rowI).getCell(1) != null
                        && sheet.getRow(rowI).getCell(2).getStringCellValue().isEmpty()) {
                    if (rowCount > 0) {
                        WriteMergeCell(workbook, rowI - rowCount, rowI - 1, 1, 1, rotateArray.get(index), "#32CD32");
                        index++;
                    }
                    rowCount = 0;
                } else {
                    rowCount++;
                }
            }

            // col 2
            rowCount = 0;
            int indexCount = 0;
            for (int rowI = 9; rowI < indexRow - 1; rowI++) {

                boolean currentIsNull = sheet.getRow(rowI).getCell(2) == null;
                boolean rightIsEmpty = sheet.getRow(rowI).getCell(3).getStringCellValue().isEmpty();
                if (!currentIsNull && rightIsEmpty) {
                    if (rowCount > 0) {
                        WriteMergeCell(workbook, rowI - rowCount, rowI - 1, 2, 2, countList.get(indexCount), "#ccffff");
                        indexCount++;
                    }
                    rowCount = 0;
                } else if (currentIsNull && rightIsEmpty) {
                    continue;
                } else {
                    rowCount++;
                }
            }

            int rowNum = sheet.getLastRowNum();
            int colNum = sheet.getRow(rowNum).getLastCellNum();

            // auto size column

            for (int c = 0; c < colNum; c++) {
                sheet.autoSizeColumn(c);
            }

            // border
            for (int r = 3; r < rowNum; r++) {
                Row row = sheet.getRow(r);
                for (int c = 4; c < colNum; c++) {
                    Cell cell = row.getCell(c);
                    if (cell == null) {
                        setCell(workbook, row, c, emptyCell, "");
                    }
                }
            }

            FileOutputStream fileOut = new FileOutputStream(fileName);
            workbook.write(fileOut);
            fileOut.close();
            workbook.close();
        } catch (Exception e) {
            System.out.print(e.getMessage());
        }
    }

}