package com.createmsdocs.controller;

import com.createmsdocs.dto.MemberDTO;
import com.createmsdocs.exporter.excel.ExcelExporter;
import com.createmsdocs.exporter.excel.ExcelStyledFromDB;
import com.createmsdocs.service.MemberService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;

@Controller
@RequestMapping("/members/export/excel")
public class ExcelController {

    @Autowired
    private MemberService memberService;

    @GetMapping("/simpleFromDB")
    public void exportToExcel(HttpServletResponse response) throws IOException {
        response.setContentType("application/octet-stream");
        DateFormat dateFormatter = new SimpleDateFormat("yyyy-MM-dd_HH:mm:ss");
        String currentDateTime = dateFormatter.format(new Date());

        String headerKey = "Content-Disposition";
        String headerValue = "attachment; filename=members_" + currentDateTime + ".xlsx";
        response.setHeader(headerKey, headerValue);

        List<MemberDTO> listMembers = memberService.findAll();

        ExcelExporter excelExporter = new ExcelExporter(listMembers);

        excelExporter.export(response);
    }

    @GetMapping("/cellStyle")
    public void exportCellStyle(HttpServletResponse response) throws IOException{
        response.setContentType("application/octet-stream");
        DateFormat dateFormatter = new SimpleDateFormat("yyyy-MM-dd_HH:mm:ss");
        String currentDateTime = dateFormatter.format(new Date());

        String headerKey = "Content-Disposition";
        String headerValue = "attachment; filename=users_" + currentDateTime + ".xlsx";
        response.setHeader(headerKey, headerValue);

        ExcelExporter excelExporter = new ExcelExporter();

        excelExporter.exportCellStyle(response);
    }

    @GetMapping("/styledFromDB")
    public void exportStyledFromDB(HttpServletResponse response) throws IOException{
        response.setContentType("application/octet-stream");
        DateFormat dateFormatter = new SimpleDateFormat("yyyy-MM-dd_HH:mm:ss");
        String currentDateTime = dateFormatter.format(new Date());

        String headerKey = "Content-Disposition";
        String headerValue = "attachment; filename=members_" + currentDateTime + ".xlsx";
        response.setHeader(headerKey, headerValue);

        List<MemberDTO> listMembers = memberService.findAll();
        ExcelStyledFromDB exporter = new ExcelStyledFromDB(listMembers);

        exporter.export(response);
    }
}
