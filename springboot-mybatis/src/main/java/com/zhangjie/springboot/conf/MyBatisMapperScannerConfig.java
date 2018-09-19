package com.zhangjie.springboot.conf;

import org.mybatis.spring.mapper.MapperScannerConfigurer;
import org.springframework.context.annotation.Bean;

/**
 * mybatis扫描配置
 * @author Administrator
 *
 */
public class MyBatisMapperScannerConfig {
	
	@Bean
	public MapperScannerConfigurer mapperScannerConfigurer(){
		MapperScannerConfigurer mapperScannerConfigurer = new MapperScannerConfigurer();
		
		mapperScannerConfigurer.setSqlSessionFactoryBeanName("sqlSessionFactory");
		
		mapperScannerConfigurer.setBasePackage("classpath:mapper/*");
		
		return mapperScannerConfigurer;
	}
}
