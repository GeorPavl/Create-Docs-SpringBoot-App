package com.createmsdocs.exporter.excel;

import com.createmsdocs.dto.MemberDTO;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.List;

public class ExcelStyledFromDB {

    private XSSFWorkbook workbook;
    private XSSFSheet sheet;
    private List<MemberDTO> listMembers;
    private Integer rowCount;

    public ExcelStyledFromDB(List<MemberDTO> listMembers) {
        this.listMembers = listMembers;
        this.workbook = new XSSFWorkbook();
        this.rowCount = 1;
    }

    public void export(HttpServletResponse response) throws IOException {
        writeTitle();
        writeSubtitle();
        writeParagraph();
        writeHeaderLine();
        writeDataLines();

        ServletOutputStream outputStream = response.getOutputStream();
        workbook.write(outputStream);
        workbook.close();

        outputStream.close();
    }

    private void writeTitle(){
        sheet = workbook.createSheet("styledFromDB");
        Row row = sheet.createRow(rowCount);

        CellStyle styleTitle = workbook.createCellStyle();

        XSSFFont fontTitle = workbook.createFont();
        fontTitle.setBold(true);
        fontTitle.setFontHeight(24);
        fontTitle.setColor(IndexedColors.BLACK.getIndex());
        styleTitle.setFont(fontTitle);


        styleTitle.setAlignment(HorizontalAlignment.CENTER);
        styleTitle.setFillPattern(FillPatternType.FINE_DOTS);
        styleTitle.setFillBackgroundColor(IndexedColors.LIGHT_YELLOW.getIndex());

        sheet.addMergedRegion(
                new CellRangeAddress(
                        1, //first row (0-based)
                        3, //last row (0-based)
                        0, //first column (0-based)
                        14 //last column (0-based)
                )
        );

        createCell(row,0,"Export .XLS files with Apache POI", styleTitle);
        rowCount += 3;
    }

    private void writeSubtitle(){
        rowCount++;

        CellStyle styleSubtitle = workbook.createCellStyle();

        XSSFFont fontSubtitle = workbook.createFont();
        fontSubtitle.setFontHeight(18);
        fontSubtitle.setUnderline(FontUnderline.SINGLE);
        styleSubtitle.setFont(fontSubtitle);

        Row rowFirst = sheet.createRow(rowCount);
        rowFirst.setHeight((short) 400);

        createCell(rowFirst,0,"Created by: Geor Pavl",styleSubtitle);

        rowCount += 2;

        Row rowSecond = sheet.createRow(rowCount);
        rowSecond.setHeight((short) 400);

        createCell(rowSecond,0, "Using: Spring Boot, JPA, Hibernate, MySQL, Apache POI, MySQL.",styleSubtitle);

        rowCount++;

        sheet.addMergedRegion(
                new CellRangeAddress(
                        rowFirst.getRowNum(), //first row (0-based)
                        rowFirst.getRowNum(), //last row (0-based)
                        0, //first column (0-based)
                        9 //last column (0-based)
                )
        );

        sheet.addMergedRegion(
                new CellRangeAddress(
                        rowSecond.getRowNum(), //first row (0-based)
                        rowSecond.getRowNum(), //last row (0-based)
                        0, //first column (0-based)
                        9 //last column (0-based)
                )
        );
    }

    private void writeParagraph(){
        rowCount++;

        CellStyle styleParagraph = workbook.createCellStyle();

        XSSFFont fontSubtitle = workbook.createFont();
        fontSubtitle.setFontHeight(14);
        styleParagraph.setFont(fontSubtitle);
        styleParagraph.setAlignment(HorizontalAlignment.JUSTIFY);
        styleParagraph.setVerticalAlignment(VerticalAlignment.JUSTIFY);

        Row row = sheet.createRow(rowCount);
        row.setHeight((short) 800);

        String myParagraph = "Apache POI, a project run by the Apache Software Foundation, and previously a sub-project of the Jakarta Project, provides pure Java libraries for reading and writing files in Microsoft Office formats, such as Word, PowerPoint and Excel.";
        createCell(row,0,myParagraph,styleParagraph);

        sheet.addMergedRegion(
                new CellRangeAddress(
                        row.getRowNum(), //first row (0-based)
                        row.getRowNum(), //last row (0-based)
                        0, //first column (0-based)
                        14 //last column (0-based)
                )
        );
        rowCount += 2;
    }

    private void writeHeaderLine() {
        Row row = sheet.createRow(rowCount++);

        CellStyle style = workbook.createCellStyle();
        XSSFFont font = workbook.createFont();

        font.setBold(true);
        font.setFontHeight(16);
        style.setFont(font);
        style.setFillBackgroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
        style.setFillPattern(FillPatternType.FINE_DOTS);

        createCell(row, 0, "Member ID", style);
        createCell(row, 1, "First Name", style);
        createCell(row, 2, "Last Name", style);
        createCell(row, 3, "E-mail", style);
    }

    private void writeDataLines() {
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


}
