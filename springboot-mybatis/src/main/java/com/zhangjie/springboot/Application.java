package com.zhangjie.springboot;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.context.annotation.ComponentScan;
import org.springframework.transaction.annotation.EnableTransactionManagement;

/**   
 * @Description: 启动类
 */
@SpringBootApplication 
@EnableTransactionManagement
@ComponentScan
public class Application {
	
	public static void main(String[] args) {
		System.setProperty("server.port", "8080"); //do config springboot server port
		SpringApplication.run(Application.class, args);
	  }
}
