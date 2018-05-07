import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

public class App {
    public static void main(String... args) throws IOException {

        Workbook workbook = new XSSFWorkbook();

        Sheet sheet = workbook.createSheet("Test Sheet");

        // change default width
        sheet.setColumnWidth(0, 6000);
        sheet.setColumnWidth(1, 4000);

        // styles
        CellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        // font styles
        XSSFFont font = ((XSSFWorkbook) workbook).createFont();
        font.setFontName("Arial");
        font.setFontHeightInPoints((short) 16);
        font.setBold(true);

        headerStyle.setFont(font);
        headerStyle.setAlignment(HorizontalAlignment.CENTER);

        // first row
        Row header = sheet.createRow(0);
        // first column
        Cell headerCell = header.createCell(0);
        headerCell.setCellValue("Name");
        headerCell.setCellStyle(headerStyle);
        // second column
        headerCell = header.createCell(1);
        headerCell.setCellValue("Age");
        headerCell.setCellStyle(headerStyle);

        // cell style
        CellStyle style = workbook.createCellStyle();
        style.setWrapText(true);

        // content - second row
        Row row = sheet.createRow(1);
        // Line 1 Column 0
        Cell cell = row.createCell(0);
        cell.setCellValue("John Smith");
        cell.setCellStyle(style);
        // Line 1 Column 1
        cell = row.createCell(1);
        cell.setCellValue(20);
        cell.setCellStyle(style);

        // Row 3 merge 10 cells
        CellRangeAddress region = new CellRangeAddress( 3, 3, 0, 10);
        sheet.addMergedRegion(region);

        // Row 3 content
        Row row1 = sheet.createRow(3);
        // set height row
        row1.setHeight((short)(sheet.getDefaultRowHeight() * 1.5));
        headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        // Line 3 column 0
        Cell cell1 = row1.createCell(0);
        cell1.setCellValue("Test merge");
        cell1.setCellStyle(headerStyle);

        // Generate file
        File currDir = new File(".");
        String path = currDir.getAbsolutePath();
        String fileLocation = path.substring(0, path.length() - 1) + "temp.xlsx";

        FileOutputStream outputStream = new FileOutputStream(fileLocation);
        workbook.write(outputStream);
        workbook.close();

    }

}

