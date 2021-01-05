package com.createmsdocs.controller;

import com.createmsdocs.exporter.csv.ExporterCSV;
import com.createmsdocs.service.MemberService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;

@Controller
@RequestMapping("/members/export/csv")
public class CSVController {

    @Autowired
    private MemberService memberService;

    @GetMapping("/exportCSV")
    public void exportCSVFromDB(HttpServletResponse response) throws IOException {
        ExporterCSV exporter = new ExporterCSV(memberService.findAll());
        exporter.export(response);
    }
}
