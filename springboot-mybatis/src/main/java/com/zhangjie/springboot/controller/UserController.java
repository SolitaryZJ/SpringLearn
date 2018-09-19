package com.zhangjie.springboot.controller;

import java.util.List;

import javax.annotation.Resource;

import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;

import com.zhangjie.springboot.domain.Account;
import com.zhangjie.springboot.service.UserService;

@Controller
@RequestMapping("/user")
public class UserController {
	
	@Resource(name = "userService")
	UserService userService;

	@RequestMapping(value="/getlist", method = RequestMethod.POST)
	@ResponseBody
	public List<Account> getlist(){
		return userService.getList();
	}
	
	@RequestMapping(value = "/login.htm", method = RequestMethod.GET)
	public String helloWorld(){
		return "login";
	}
 }
