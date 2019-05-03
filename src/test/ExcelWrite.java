/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package test;
  import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

/**
 *
 * @author Andrew
 */
public class ExcelWrite {
 


    private static final String FILE_NAME = "/home/mspace/Pictures/MyFirstExcel.xlsx";

    public static void main(String[] args) {

        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = allocategenerateSheet(workbook, "MydataTest_1");
        Object[][] datatypes = {
                {"+25470****010", "Si****yu", "Kenya"},
                {"+25470****010", "A****ew", "Kenya"},
                {"+25470****010", "M****te", "Kenya"},
                {"+25470****010", "E****ne", "Kenya"},
                {"+25470****010", "I****ga", "Kenya"}
        };

        int rowNum = 2;
        System.out.println("Creating excel");

        for (Object[] datatype : datatypes) {
            int sheetCounter = 1;
                if (rowNum % 65535 == 0) {
                    sheetCounter++;
                    String new_sheetName = "MydataTest_" + sheetCounter;
                    sheet = allocategenerateSheet(workbook, new_sheetName);
                    System.out.println("Creating new sheet Name: " + new_sheetName);
                    rowNum = 2;
                }
            Row row = sheet.createRow(rowNum);
            int colNum = 0;
            for (Object field : datatype) {
                Cell cell = row.createCell(colNum);
                if (field instanceof String) {
                    cell.setCellValue((String) field);
                } else if (field instanceof Integer) {
                    cell.setCellValue((Integer) field);
                }
                colNum++;
            }
            rowNum++;
        }

        try {
            FileOutputStream outputStream = new FileOutputStream(FILE_NAME);
            workbook.write(outputStream);
            workbook.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

        System.out.println("Done");
    } 
    
     private static HSSFSheet allocategenerateSheet(HSSFWorkbook wb, String sheetName) {
        HSSFSheet sheet = wb.createSheet(sheetName);
        /*You can add style here too or write header here*/
        Map<String, CellStyle> styles = createStyles(wb);

        PrintSetup printSetup = sheet.getPrintSetup();
        printSetup.setLandscape(true);
        sheet.setFitToPage(true);
        sheet.setHorizontallyCenter(true);

        //title row
        Row titleRow = sheet.createRow(0);
        titleRow.setHeightInPoints(45);
        Cell titleCell = titleRow.createCell(0);
        titleCell.setCellValue("Testing Report Generation");
        titleCell.setCellStyle(styles.get("title"));
        sheet.addMergedRegion(CellRangeAddress.valueOf("$A$1:$C$1"));

        String[] titles = {"Mobile Number", "Name", "Country"};

        HSSFRow row = sheet.createRow(1);
        row.setHeightInPoints(40);

        Cell headerCell;
        for (int i = 0; i < titles.length; i++) {
            headerCell = row.createCell(i);
            headerCell.setCellValue(titles[i]);
            headerCell.setCellStyle(styles.get("header"));
        }

        return sheet;
    }
      private static Map<String, CellStyle> createStyles(Workbook wb) {

        Map<String, CellStyle> styles = new HashMap<>();
        CellStyle style;
        Font titleFont = wb.createFont();
        titleFont.setFontHeightInPoints((short) 18);
        titleFont.setBoldweight(Font.BOLDWEIGHT_BOLD);
        style = wb.createCellStyle();
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        style.setFont(titleFont);
        styles.put("title", style);

        Font monthFont = wb.createFont();
        monthFont.setFontHeightInPoints((short) 11);
        monthFont.setColor(IndexedColors.WHITE.getIndex());
        style = wb.createCellStyle();
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        style.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        style.setFont(monthFont);
        style.setWrapText(true);
        styles.put("header", style);

        style = wb.createCellStyle();
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setWrapText(true);
        style.setBorderRight(CellStyle.BORDER_THIN);
        style.setRightBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderLeft(CellStyle.BORDER_THIN);
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderTop(CellStyle.BORDER_THIN);
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderBottom(CellStyle.BORDER_THIN);
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        styles.put("cell", style);

        style = wb.createCellStyle();
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        style.setDataFormat(wb.createDataFormat().getFormat("0.00"));
        styles.put("formula", style);

        style = wb.createCellStyle();
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        style.setFillForegroundColor(IndexedColors.GREY_40_PERCENT.getIndex());
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        style.setDataFormat(wb.createDataFormat().getFormat("0.00"));
        styles.put("formula_2", style);

        return styles;
    }
//for (OPTOut ret : listhere) {
//                int sheetCounter = 1;
//                if (rowNum % 65535 == 0) {
//                    sheetCounter++;
//                    String new_sheetName = "mysheet_" + sheetCounter;
//                    sheet = allocategenerateOPTXSLSheet(wb, new_sheetName);
//                    System.out.println("Creating new sheet Name: " + new_sheetName);
//                    rowNum = 2;
//                }
//                HSSFRow row = sheet.createRow(rowNum);
//                row.createCell(0).setCellValue(ret.get*****());
//                row.createCell(1).setCellValue(ret.get*****());
//                row.createCell(2).setCellValue(ret.get*****());
//                row.createCell(3).setCellValue(ret.get*****());
//                row.createCell(4).setCellValue(ret.get*****());
//                row.createCell(5).setCellValue(ret.get*****());
//
//                row.createCell(6).setCellValue(ret.get*****());
//                rowNum++;
//            }
}
