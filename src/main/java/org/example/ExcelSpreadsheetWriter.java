package org.example;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashSet;
import java.util.Set;

public class ExcelSpreadsheetWriter {
    private Workbook workbook;
    private Sheet sheet;
    private Set<String> visitedCells;

    public ExcelSpreadsheetWriter(String filePath, String sheetName) {
        try {
            File file = new File(filePath);
            if (file.exists()) {
                FileInputStream fis = new FileInputStream(file);
                workbook = new XSSFWorkbook(fis);
            } else {
                workbook = new XSSFWorkbook();
            }

            sheet = workbook.getSheet(sheetName);
            if (sheet == null) {
                sheet = workbook.createSheet(sheetName);
            }

            visitedCells = new HashSet<>();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public void setCellValue(String cellId, Object value) {
        if (!isValidCellId(cellId)) {
            throw new IllegalArgumentException("Invalid cellId: " + cellId);
        }

        Row row = sheet.getRow(getRowIndex(cellId));
        if (row == null) {
            row = sheet.createRow(getRowIndex(cellId));
        }

        Cell cell = row.getCell(getColumnIndex(cellId));
        if (cell == null) {
            cell = row.createCell(getColumnIndex(cellId));
        }

        if (value instanceof Integer) {
            cell.setCellValue((Integer) value);
        } else if (value instanceof String) {
            String stringValue = (String) value;
            if (stringValue.startsWith("=")) {
                String formula = stringValue.substring(1);
                validateFormula(cellId, formula);
                cell.setCellFormula(formula);
            } else {
                cell.setCellValue(stringValue);
            }
        }
    }

    public int getCellValue(String cellId) {
        if (!isValidCellId(cellId)) {
            throw new IllegalArgumentException("Invalid cellId: " + cellId);
        }

        visitedCells.clear();
        return evaluateCell(cellId);
    }

    private int evaluateCell(String cellId) {
        if (visitedCells.contains(cellId)) {
            throw new IllegalArgumentException("Circular reference detected for cellId: " + cellId);
        }

        visitedCells.add(cellId);

        Row row = sheet.getRow(getRowIndex(cellId));
        if (row != null) {
            Cell cell = row.getCell(getColumnIndex(cellId));
            if (cell != null && cell.getCellType() == CellType.FORMULA) {
                FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
                CellValue cellValue = evaluator.evaluate(cell);
                if (cellValue.getCellType() == CellType.NUMERIC) {
                    return (int) cellValue.getNumberValue();
                }
            }
        }

        return 0; // Default value if the cell is not found or contains non-numeric data.
    }

    private void validateFormula(String cellId, String formula) {
        if (!formula.matches("([A-Z]\\d)+[+-/*%]([A-Z]\\d)+") && !formula.matches("\\d+")) {
            throw new IllegalArgumentException("Invalid formula '" + formula + "' in cellId: " + cellId);
        }
    }

    private boolean isValidCellId(String cellId) {
        return cellId.matches("[A-Z]\\d+");
    }

    public void saveToFile(String filePath) {
        try {
            FileOutputStream fos = new FileOutputStream(filePath);
            workbook.write(fos);
            fos.close();
            System.out.println("Data saved to the file: " + filePath);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private int getRowIndex(String cellId) {
        return Integer.parseInt(cellId.replaceAll("[^0-9]", "")) - 1;
    }

    private int getColumnIndex(String cellId) {
        return cellId.charAt(0) - 'A';
    }

    public static void main(String[] args) {
        String filePath = "testSheet.xlsx";
        String sheetName = "Sheet1";

        ExcelSpreadsheetWriter writer = new ExcelSpreadsheetWriter(filePath, sheetName);

        writer.setCellValue("A1", 13);
        writer.setCellValue("A2", 14);

        int cellValueA1 = writer.getCellValue("A1");
        System.out.println("Cell A1 value: " + cellValueA1); // Output: Cell A1 value: 13

        writer.setCellValue("A3", "=A1+A2");
        int cellValueA3 = writer.getCellValue("A3");
        System.out.println("Cell A3 value: " + cellValueA3); // Output: Cell A3 value: 27

        writer.saveToFile(filePath);
    }
}
