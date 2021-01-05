package com.createmsdocs.exporter.word;

import com.createmsdocs.dto.MemberDTO;
import org.apache.poi.ss.usermodel.FontFamily;
import org.apache.poi.xwpf.usermodel.*;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;

public class WordExporterFromDB {

    private XWPFDocument document;
    private FileOutputStream out;
    private List<MemberDTO> list;

    public WordExporterFromDB(List<MemberDTO> memberDTOS) throws FileNotFoundException {
        document = new XWPFDocument();
        out = new FileOutputStream(new File("wordFromDB.docx"));
        list = memberDTOS;
    }

    public void export(HttpServletResponse response) throws IOException {
        response.setContentType("application/octet-stream");
        DateFormat dateFormatter = new SimpleDateFormat("yyyy-MM-dd_HH:mm:ss");
        String currentDateTime = dateFormatter.format(new Date());

        String headerKey = "Content-Disposition";
        String headerValue = "attachment; filename=members_" + currentDateTime + ".docx";
        response.setHeader(headerKey, headerValue);

        writeTitle();
        writeSubtitle();
        writeSimpleParagraph();
        writeTable();
        writeAlignedParagraph();

        ServletOutputStream outputStream = response.getOutputStream();
        document.write(outputStream);
        document.close();

        outputStream.close();
    }

    private void writeTitle(){
        XWPFParagraph title = document.createParagraph();
        title.setAlignment(ParagraphAlignment.CENTER);

        XWPFRun run = title.createRun();
        run.addBreak();
        run.setBold(true);
        run.setFontSize(24);
        run.setFontFamily(FontFamily.ROMAN.name());
        run.setText("Export Word Files");
        run.addBreak();
    }

    private void writeSubtitle(){
        XWPFParagraph subtitle = document.createParagraph();
        subtitle.setAlignment(ParagraphAlignment.CENTER);

        XWPFRun run = subtitle.createRun();
        run.setItalic(true);
        run.setFontSize(18);
        run.setFontFamily(FontFamily.SCRIPT.name());
        run.setText("Using Spring Boot, Apache POI, JPA, Hibernate. Thymeleaf");
        run.addBreak();
    }

    private void writeSimpleParagraph(){
        XWPFParagraph paragraph = document.createParagraph();

        XWPFRun run = paragraph.createRun();
        run.setFontSize(12);
        run.setFontFamily(FontFamily.MODERN.name());
        run.setText("Lorem Ipsum is simply dummy text of the printing and typesetting industry.");
        run.addBreak();

        paragraph.setSpacingAfter(40);
    }

    private void writeTable(){
        //create table
        XWPFTable table = document.createTable();
        int rowCount = 0;

        //create first row
        XWPFTableRow firstRow = table.getRow(rowCount);
        firstRow.getCell(0).setText("First Name");
        firstRow.addNewTableCell().setText("Last Name");
        firstRow.addNewTableCell().setText("Email");

        for (MemberDTO tempMember : list){
            XWPFTableRow tempRow = table.createRow();
            tempRow.getCell(0).setText(tempMember.getFirstName());
            tempRow.getCell(1).setText(tempMember.getLastName());
            tempRow.getCell(2).setText(tempMember.getEmail());
        }
    }

    private void writeAlignedParagraph(){
        XWPFParagraph paragraph = document.createParagraph();
        paragraph.setAlignment(ParagraphAlignment.RIGHT);

        XWPFRun run = paragraph.createRun();
        run.addBreak();
        run.setItalic(true);
        run.setFontSize(14);
        run.setColor("990000");
        run.setText("Created in 2021");
    }
}
