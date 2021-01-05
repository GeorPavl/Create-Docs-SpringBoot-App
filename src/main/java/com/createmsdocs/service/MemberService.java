package com.createmsdocs.service;

import com.createmsdocs.dto.MemberDTO;
import com.createmsdocs.entity.Member;

import java.util.List;

public interface MemberService {

    Member dtoToEntity(MemberDTO source);

    MemberDTO writeMemberDTO(Member source);

    List<MemberDTO> findAll();
}
