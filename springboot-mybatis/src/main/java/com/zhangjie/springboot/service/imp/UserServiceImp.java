package com.zhangjie.springboot.service.imp;

import java.util.List;
import java.util.Map;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import com.zhangjie.springboot.dao.UserDao;
import com.zhangjie.springboot.domain.Account;
import com.zhangjie.springboot.service.UserService;

@Service("userService")
public class UserServiceImp implements UserService{
	
	@Autowired
	UserDao userDao;
	
	@Override
	public List<Account> getList() {
		return userDao.getListAccount();
	}

}
