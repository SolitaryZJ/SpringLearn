package com.zhangjie.springboot.shrio;

import org.apache.shiro.session.Session;
import org.apache.shiro.session.SessionListener;

public class ShiroSessionListener implements SessionListener{

	@Override
	public void onStart(Session session) {
		System.out.println("创建会话:"+session.getId());
	}

	@Override
	public void onStop(Session session) {
		System.out.println("创建停止:"+session.getId());
	}

	@Override
	public void onExpiration(Session session) {
		System.out.println("创建过期:"+session.getId());
	}


}
