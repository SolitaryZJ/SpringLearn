package com.zhangjie.springboot.controller;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.List;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.converter.ExcelToHtmlConverter;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hwpf.usermodel.Picture;
import org.w3c.dom.Document;

public class PoiExcelToHtml {

	public static String getExeclToHtml(File file) throws Exception {
		InputStream input = new FileInputStream(file);

		HSSFWorkbook excelBook = new HSSFWorkbook(input);
		ExcelToHtmlConverter excelToHtmlConverter = new ExcelToHtmlConverter(
				DocumentBuilderFactory.newInstance().newDocumentBuilder()
						.newDocument());
		// excelToHtmlConverter.setUseDivsToSpan(true); //使用div放置内容
		excelToHtmlConverter.setOutputColumnHeaders(false); // 头部序号标签
		excelToHtmlConverter.setOutputRowNumbers(false); // 行号
		excelToHtmlConverter.processWorkbook(excelBook);

		List pics = excelBook.getAllPictures();
		if (pics != null) {
			for (int i = 0; i < pics.size(); i++) {
				Picture pic = (Picture) pics.get(i);
				try {
					pic.writeImageContent(new FileOutputStream(
							StaticFinal.FILE_PATH + pic.suggestFullFileName()));
				} catch (FileNotFoundException e) {
					e.printStackTrace();
				}
			}
		}
		Document htmlDocument = excelToHtmlConverter.getDocument();
		ByteArrayOutputStream outStream = new ByteArrayOutputStream();
		DOMSource domSource = new DOMSource(htmlDocument);
		StreamResult streamResult = new StreamResult(outStream);
		TransformerFactory tf = TransformerFactory.newInstance();
		Transformer serializer = tf.newTransformer();
		serializer.setOutputProperty(OutputKeys.ENCODING, "utf-8");
		serializer.setOutputProperty(OutputKeys.INDENT, "yes");
		serializer.setOutputProperty(OutputKeys.METHOD, "html");
		serializer.transform(domSource, streamResult);
		outStream.close();
		String content = new String(outStream.toByteArray());
		return content;
	}

	final static String path = "D:\\";
	final static String file = "证券投资基金估值表_国金慧泉精选对冲1号限额特定集合资产管理计划_2014-01-06.xls";

	public static void main(String args[]) throws Exception {
		InputStream input = new FileInputStream(path + file);

		HSSFWorkbook excelBook = new HSSFWorkbook(input);
		/*ExcelToHtmlConverter excelToHtmlConverter = new ExcelToHtmlConverter(
				DocumentBuilderFactory.newInstance().newDocumentBuilder()
						.newDocument());*/
		MyConverter excelToHtmlConverter = new MyConverter(DocumentBuilderFactory.newInstance().newDocumentBuilder()
				.newDocument());
		// excelToHtmlConverter.setUseDivsToSpan(true); //使用div放置内容
		excelToHtmlConverter.setOutputColumnHeaders(false); // 头部序号标签
		excelToHtmlConverter.setOutputRowNumbers(false); // 行号
		excelToHtmlConverter.processWorkbook(excelBook);

		List pics = excelBook.getAllPictures();
		if (pics != null) {
			for (int i = 0; i < pics.size(); i++) {
				Picture pic = (Picture) pics.get(i);
				try {
					pic.writeImageContent(new FileOutputStream(path
							+ pic.suggestFullFileName()));
				} catch (FileNotFoundException e) {
					e.printStackTrace();
				}
			}
		}
		Document htmlDocument = excelToHtmlConverter.getDocument();
		ByteArrayOutputStream outStream = new ByteArrayOutputStream();
		DOMSource domSource = new DOMSource(htmlDocument);
		StreamResult streamResult = new StreamResult(outStream);
		TransformerFactory tf = TransformerFactory.newInstance();
		Transformer serializer = tf.newTransformer();
		serializer.setOutputProperty(OutputKeys.ENCODING, "utf-8");
		serializer.setOutputProperty(OutputKeys.INDENT, "yes");
		serializer.setOutputProperty(OutputKeys.METHOD, "html");
		serializer.transform(domSource, streamResult);
		outStream.close();
		String content = new String(outStream.toByteArray());

		// 加入特殊样式overflow: scroll;
		// 解析content
		String withoutNString = content.replace("\n", "");
		// 过滤\r 转换成空
		String withoutRString = withoutNString.replace("\r", "");
		// 过滤\t 转换成空
		String withoutTString = withoutRString.replace("\t", "");
		// 过滤\ 转换成空
		String newString = withoutTString.replace("\\", "");
		// 获取html中的body标签
		// String result = Regex.Match(newString,
		// @"<body.*>.*</body>").ToString();

		FileUtils.writeStringToFile(new File(path, "exportExcel3.html"),
				content, "utf-8");
		System.err.println("html生产完毕");
	}
}