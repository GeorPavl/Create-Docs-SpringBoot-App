package com.createmsdocs.controller;

import com.createmsdocs.dto.MemberDTO;
import com.createmsdocs.service.MemberService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import java.util.List;

@RestController
public class MemberController {

    @Autowired
    private MemberService memberService;

    @GetMapping("/findAll")
    public List<MemberDTO> findAll(){
        return memberService.findAll();
    }
}
