package webserviceTest.classTest;

import javax.xml.ws.Endpoint;

import webserviceTest.service.imp.IMathServiceImp;

public class TestMain {
	
	public static void main(String[] args) {
		String address = "http://localhost:8888/ns";
		Endpoint.publish(address, new IMathServiceImp());
	}
	
}
