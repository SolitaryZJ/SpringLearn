package com.zhangjie.springboot.controller;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.converter.ExcelToHtmlConverter;
import org.apache.poi.hssf.converter.ExcelToHtmlUtils;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hwpf.converter.HtmlDocumentFacade;
import org.apache.poi.ss.util.CellRangeAddress;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

public class TestExcelToHtml extends ExcelToHtmlConverter {

	private  HtmlDocumentFacade htmlDocumentFacade1;

	private String cssClassPrefixRow = "r";
	private String cssClassPrefixTable = "t";
	
	public TestExcelToHtml(Document doc) {
		super(doc);
	}

	public TestExcelToHtml(HtmlDocumentFacade htmlDocumentFacade) {
		super(htmlDocumentFacade);
		this.htmlDocumentFacade1 = htmlDocumentFacade;
	}

	protected void processSheet(HSSFSheet sheet) {
		processSheetHeader(this.htmlDocumentFacade1.getBody(), sheet);

		int physicalNumberOfRows = sheet.getPhysicalNumberOfRows();
		if (physicalNumberOfRows <= 0) {
			return;
		}
		Element table = this.htmlDocumentFacade1.createTable();
		this.htmlDocumentFacade1.addStyleClass(table, this.cssClassPrefixTable,
				"border-collapse:collapse;border-spacing:0;");

		Element tableBody = this.htmlDocumentFacade1.createTableBody();

		CellRangeAddress[][] mergedRanges = ExcelToHtmlUtils
				.buildMergedRangesMap(sheet);

		List<Element> emptyRowElements = new ArrayList(physicalNumberOfRows);
		int maxSheetColumns = 1;
		for (int r = sheet.getFirstRowNum(); r <= sheet.getLastRowNum(); r++) {
			HSSFRow row = sheet.getRow(r);

			if (row != null) {

				if ((isOutputHiddenRows()) || (!row.getZeroHeight())) {

					Element tableRowElement = this.htmlDocumentFacade1
							.createTableRow();
					this.htmlDocumentFacade1.addStyleClass(tableRowElement,
							this.cssClassPrefixRow, "height:" + row.getHeight()
									/ 20.0F + "pt;");

					int maxRowColumnNumber = processRow(mergedRanges, row,
							tableRowElement);

					if (maxRowColumnNumber == 0) {
						emptyRowElements.add(tableRowElement);
					} else {
						if (!emptyRowElements.isEmpty()) {
							for (Element emptyRowElement : emptyRowElements) {
								tableBody.appendChild(emptyRowElement);
							}
							emptyRowElements.clear();
						}

						tableBody.appendChild(tableRowElement);
					}
					maxSheetColumns = Math.max(maxSheetColumns,
							maxRowColumnNumber);
				}
			}
		}
		processColumnWidths(sheet, maxSheetColumns, table);

		if (isOutputColumnHeaders()) {
			processColumnHeaders(sheet, maxSheetColumns, table);
		}

		table.appendChild(tableBody);

		this.htmlDocumentFacade1.getBody().appendChild(table);
	}

}
