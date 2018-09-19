package webserviceTest.test;

import java.io.IOException;

import javax.xml.namespace.QName;
import javax.xml.soap.MessageFactory;
import javax.xml.soap.SOAPBody;
import javax.xml.soap.SOAPEnvelope;
import javax.xml.soap.SOAPException;
import javax.xml.soap.SOAPMessage;
import javax.xml.soap.SOAPPart;

import org.junit.Test;

public class TestSoap {

	@Test
	public void test01() throws SOAPException, IOException{
		//1、创建消息工厂
		MessageFactory newInstance = MessageFactory.newInstance();
		//2、根据消息工厂创建SoapMessage
		SOAPMessage createMessage = newInstance.createMessage();
		//3、创建soapPart
		SOAPPart soapPart = createMessage.getSOAPPart();
		//4、获得soapEnvelope
		SOAPEnvelope envelope = soapPart.getEnvelope();
		//5、可以通过soapEnvelope有效的获取响应的body和header等信息
		SOAPBody body = envelope.getBody();
		//6、根据QName创建相应的节点，QName就是一个带有命名空间的<zj:add xmlns="http://www.jie.com/webservice">
		QName qName = new QName("http://www.jie.com/webservice","add","zj");
		body.addBodyElement(qName).setValue("123123");
		//打印内容
		createMessage.writeTo(System.out);
	}
}
