package com.createmsdocs.controller;

import com.createmsdocs.dto.MemberDTO;
import com.createmsdocs.exporter.word.WordExporter;
import com.createmsdocs.exporter.word.WordExporterFromDB;
import com.createmsdocs.service.MemberService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.List;

@Controller
@RequestMapping("/members/export/word")
public class WordController {

    @Autowired
    private MemberService memberService;

    @GetMapping("/simpleWord")
    public void exportSimpleWord(HttpServletResponse response) throws IOException {
        WordExporter wordExporter = new WordExporter();
        wordExporter.export(response);
    }

    @GetMapping("/wordFromDB")
    public void exportWordFromDB(HttpServletResponse response) throws IOException {
        List<MemberDTO> list = memberService.findAll();
        WordExporterFromDB exporter = new WordExporterFromDB(list);
        exporter.export(response);
    }
}
