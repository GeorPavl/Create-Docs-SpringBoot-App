package com.createmsdocs.exporter.excel;

import com.createmsdocs.dto.MemberDTO;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.List;

public class ExcelExporter {
    private XSSFWorkbook workbook;
    private XSSFSheet sheet;
    private List<MemberDTO> listMembers;

    public ExcelExporter(List<MemberDTO> listMembers) {
        this.listMembers = listMembers;
        workbook = new XSSFWorkbook();
    }

    public ExcelExporter(){
        workbook = new XSSFWorkbook();
    }

    public void export(HttpServletResponse response) throws IOException {
        writeHeaderLine();
        writeDataLines();

        ServletOutputStream outputStream = response.getOutputStream();
        workbook.write(outputStream);
        workbook.close();

        outputStream.close();
    }

    private void writeHeaderLine() {
        sheet = workbook.createSheet("Members");

        Row row = sheet.createRow(0);

        CellStyle style = workbook.createCellStyle();
        XSSFFont font = workbook.createFont();
        font.setBold(true);
        font.setFontHeight(16);
        style.setFont(font);

        createCell(row, 0, "Member ID", style);
        createCell(row, 1, "First Name", style);
        createCell(row, 2, "Last Name", style);
        createCell(row, 3, "E-mail", style);
    }

    private void createCell(Row row, int columnCount, Object value, CellStyle style) {
        sheet.autoSizeColumn(columnCount);
        Cell cell = row.createCell(columnCount);
        if (value instanceof Integer) {
            cell.setCellValue((Integer) value);
        } else if (value instanceof Boolean) {
            cell.setCellValue((Boolean) value);
        }else if (value instanceof Long){
            cell.setCellValue((Long) value);
        }else {
            cell.setCellValue((String) value);
        }
        cell.setCellStyle(style);
    }

    private void writeDataLines() {
        int rowCount = 1;

        CellStyle style = workbook.createCellStyle();
        XSSFFont font = workbook.createFont();
        font.setFontHeight(14);
        style.setFont(font);

        for (MemberDTO member : listMembers) {
            Row row = sheet.createRow(rowCount++);
            int columnCount = 0;

            createCell(row, columnCount++, member.getId(), style);
            createCell(row, columnCount++, member.getFirstName(), style);
            createCell(row, columnCount++, member.getLastName(), style);
            createCell(row, columnCount++, member.getEmail(), style);
        }
    }

    public void exportCellStyle(HttpServletResponse response) throws IOException {
        XSSFSheet spreadsheet = workbook.createSheet("cellstyle");
        XSSFRow row = spreadsheet.createRow((short) 1);
        row.setHeight((short) 800);
        XSSFCell cell = (XSSFCell) row.createCell((short) 1);
        cell.setCellValue("test of merging");
        //MERGING CELLS
        //this statement for merging cells

        spreadsheet.addMergedRegion(
                new CellRangeAddress(
                        1, //first row (0-based)
                        1, //last row (0-based)
                        1, //first column (0-based)
                        4 //last column (0-based)
                )
        );

        //CELL Alignment
        row = spreadsheet.createRow(5);
        cell = (XSSFCell) row.createCell(0);
        row.setHeight((short) 800);

        // Top Left alignment
        XSSFCellStyle style1 = workbook.createCellStyle();
        spreadsheet.setColumnWidth(0, 8000);
        style1.setAlignment(HorizontalAlignment.LEFT);
        style1.setVerticalAlignment(VerticalAlignment.TOP);
        cell.setCellValue("Top Left");
        cell.setCellStyle(style1);
        row = spreadsheet.createRow(6);
        cell = (XSSFCell) row.createCell(1);
        row.setHeight((short) 800);

        // Center Align Cell Contents
        XSSFCellStyle style2 = workbook.createCellStyle();
        style2.setAlignment(HorizontalAlignment.CENTER);
        style2.setVerticalAlignment(VerticalAlignment.CENTER);
        cell.setCellValue("Center Aligned");
        cell.setCellStyle(style2);
        row = spreadsheet.createRow(7);
        cell = (XSSFCell) row.createCell(2);
        row.setHeight((short) 800);

        // Bottom Right alignment
        XSSFCellStyle style3 = workbook.createCellStyle();
        style3.setAlignment(HorizontalAlignment.RIGHT);
        style3.setVerticalAlignment(VerticalAlignment.BOTTOM);
        cell.setCellValue("Bottom Right");
        cell.setCellStyle(style3);
        row = spreadsheet.createRow(8);
        cell = (XSSFCell) row.createCell(3);

        // Justified Alignment
        XSSFCellStyle style4 = workbook.createCellStyle();
        style4.setAlignment(HorizontalAlignment.JUSTIFY);
        style4.setVerticalAlignment(VerticalAlignment.JUSTIFY);
        cell.setCellValue("Contents are Justified in Alignment");
        cell.setCellStyle(style4);

        //CELL BORDER
        row = spreadsheet.createRow((short) 10);
        row.setHeight((short) 800);
        cell = (XSSFCell) row.createCell((short) 1);
        cell.setCellValue("BORDER");

        XSSFCellStyle style5 = workbook.createCellStyle();
        style5.setBorderBottom(BorderStyle.THICK);
        style5.setBottomBorderColor(IndexedColors.BLUE.getIndex());
        style5.setBorderLeft(BorderStyle.DOUBLE);
        style5.setLeftBorderColor(IndexedColors.GREEN.getIndex());
        style5.setBorderRight(BorderStyle.HAIR);
        style5.setRightBorderColor(IndexedColors.RED.getIndex());
        style5.setBorderTop(BorderStyle.DASHED);
        style5.setTopBorderColor(IndexedColors.CORAL.getIndex());
        cell.setCellStyle(style5);

        //Fill Colors
        //background color
        row = spreadsheet.createRow((short) 10 );
        cell = (XSSFCell) row.createCell((short) 1);

        XSSFCellStyle style6 = workbook.createCellStyle();
        style6.setFillBackgroundColor(IndexedColors.LEMON_CHIFFON.getIndex());
        style6.setFillPattern(FillPatternType.FINE_DOTS);
        style6.setAlignment(HorizontalAlignment.FILL);
        spreadsheet.setColumnWidth(1,8000);
        cell.setCellValue("FILL BACKGROUND/FILL PATTERN");
        cell.setCellStyle(style6);

        //Foreground color
        row = spreadsheet.createRow((short) 12);
        cell = (XSSFCell) row.createCell((short) 1);

        XSSFCellStyle style7 = workbook.createCellStyle();
        style7.setFillForegroundColor(IndexedColors.BLUE.getIndex());
        style7.setFillPattern(FillPatternType.LESS_DOTS);
        style7.setAlignment(HorizontalAlignment.FILL);
        cell.setCellValue("FILL FOREGROUND/FILL PATTERN");
        cell.setCellStyle(style7);

        ServletOutputStream outputStream = response.getOutputStream();
        workbook.write(outputStream);
        workbook.close();

        outputStream.close();
    }
}
