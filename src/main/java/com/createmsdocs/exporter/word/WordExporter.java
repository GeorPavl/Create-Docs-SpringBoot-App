package com.createmsdocs.exporter.word;

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

public class WordExporter {

    private XWPFDocument document;
    private FileOutputStream out;

    public WordExporter() throws FileNotFoundException {
        document = new XWPFDocument();
        out = new FileOutputStream(new File("simpleStyledWord.docx"));
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
        run.setText("Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled it to make a type specimen book. It has survived not only five centuries, but also the leap into electronic typesetting, remaining essentially unchanged. It was popularised in the 1960s with the release of Letraset sheets containing Lorem Ipsum passages, and more recently with desktop publishing software like Aldus PageMaker including versions of Lorem Ipsum.");
        run.addBreak();

        paragraph.setSpacingAfter(40);
    }

    private void writeTable(){
        //create table
        XWPFTable table = document.createTable();

        //create first row
        XWPFTableRow tableRowOne = table.getRow(0);
        tableRowOne.getCell(0).setText("col one, row one");
        tableRowOne.addNewTableCell().setText("col two, row one");
        tableRowOne.addNewTableCell().setText("col three, row one");

        //create second row
        XWPFTableRow tableRowTwo = table.createRow();
        tableRowTwo.getCell(0).setText("col one, row two");
        tableRowTwo.getCell(1).setText("col two, row two");
        tableRowTwo.getCell(2).setText("col three, row two");

        //create third row
        XWPFTableRow tableRowThree = table.createRow();
        tableRowThree.getCell(0).setText("col one, row three");
        tableRowThree.getCell(1).setText("col two, row three");
        tableRowThree.getCell(2).setText("col three, row three");
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
