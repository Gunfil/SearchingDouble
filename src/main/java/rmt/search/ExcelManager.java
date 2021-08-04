package rmt.search;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.List;

public class ExcelManager {

    private static final ExcelManager em = new ExcelManager();
    private HSSFWorkbook book;
    private HSSFSheet sheet;

    private ExcelManager() {}

    public static ExcelManager getExcelManager() {
        return em;
    }

    public void createNewBook() {
        book = new HSSFWorkbook();
    }

    public void createSheet(String name) {
        sheet = book.createSheet(name);
    }

    public void openExcelForEdit(String file, int sheetNumber) {
        try {
            FileInputStream is = new FileInputStream(file);
            book = new HSSFWorkbook(is);
            sheet = book.getSheetAt(sheetNumber);
            is.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public void setHeader(String [] header) {
        HSSFRow headerRow = sheet.createRow(0);
        CellStyle headerStyle = book.createCellStyle();
        HSSFFont font = book.createFont();
        font.setBold(true);
        headerStyle.setFont(font);

        for (int i = 0; i < header.length; i++) {
            HSSFCell head = headerRow.createCell(i);
            head.setCellValue(header[i]);
            head.setCellStyle(headerStyle);
        }
    }

    public void write(List<String> values) {
        int rowNumber = sheet.getPhysicalNumberOfRows();
        HSSFRow row = sheet.createRow(rowNumber);
        HSSFCell cell;
        for (int i = 0; i < values.size(); i++) {
            cell = row.createCell(i);
            cell.setCellValue(values.get(i));
        }
    }

    public void save(String file) {
        try {
            FileOutputStream os = new FileOutputStream(file);
            book.write(os);
            os.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public void emptyRow() {
        int count = sheet.getPhysicalNumberOfRows();
        sheet.createRow(count);
    }

    public Sheet getSheet(File file) {
        String name = file.getName();
        if (name.endsWith(".xlsx")) {
            XSSFWorkbook workbook = null;
            try {
                workbook = new XSSFWorkbook(file);

            } catch (InvalidFormatException | IOException e) {
                e.printStackTrace();
            }
            return workbook.getSheetAt(0);
        } else {
            HSSFWorkbook workbook = null;
            try {
                FileInputStream is = new FileInputStream(file);
                workbook = new HSSFWorkbook(is);
                is.close();
            } catch (IOException e) {
                e.printStackTrace();
            }

            return workbook.getSheetAt(0);
        }
    }

    public void closeBook() {
        try {
            book.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
