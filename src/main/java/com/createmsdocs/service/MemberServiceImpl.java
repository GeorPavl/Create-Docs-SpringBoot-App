package com.createmsdocs.service;

import com.createmsdocs.dto.MemberDTO;
import com.createmsdocs.entity.Member;
import com.createmsdocs.repository.MemberRepository;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.util.ArrayList;
import java.util.List;

@Service
public class MemberServiceImpl implements MemberService{

    @Autowired
    private MemberRepository memberRepository;

    @Override
    public Member dtoToEntity(MemberDTO source) {
        Member result = new Member();
        if (source.getId() != null){
            result.setId(source.getId());
        }
        if (source.getFirstName() != null){
            result.setFirstName(source.getFirstName());
        }
        if (source.getLastName() != null){
            result.setLastName(source.getLastName());
        }
        if (source.getEmail() != null){
            result.setEmail(source.getEmail());
        }
        return result;
    }

    @Override
    public MemberDTO writeMemberDTO(Member source) {
        MemberDTO result = new MemberDTO();
        result.setId(source.getId());
        result.setFirstName(source.getFirstName());
        result.setLastName(source.getLastName());
        result.setEmail(source.getEmail());
        return result;
    }

    @Override
    public List<MemberDTO> findAll() {
        List<MemberDTO> memberDTOS = new ArrayList<>();
        for (Member tempMember : memberRepository.findAll()){
            memberDTOS.add(writeMemberDTO(tempMember));
        }
        return memberDTOS;
    }
}
