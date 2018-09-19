package com.zhangjie.springboot.dao;

import java.util.List;

import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Select;

import com.zhangjie.springboot.domain.Account;

@Mapper
public interface UserDao {

	@Select("select * from account ")
	public List<Account> getListAccount();
}
