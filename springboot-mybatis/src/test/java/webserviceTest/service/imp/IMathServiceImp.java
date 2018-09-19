package webserviceTest.service.imp;

import javax.jws.WebService;

import webserviceTest.service.MathService;

@WebService(endpointInterface="webserviceTest.service.MathService")
public class IMathServiceImp implements MathService{

	@Override
	public int add(int a, int b) {
		System.out.println("加法运算："+a+"+"+b+"="+(a+b));
		return (a+b);
	}

	@Override
	public int minus(int a, int b) {
		System.out.println("减法运算："+a+"-"+b+"="+(a-b));
		return (a-b);
	}

}
