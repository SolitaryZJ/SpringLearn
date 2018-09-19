package com.zhangjie.springboot.controller;

import java.io.File;
import java.io.FileWriter;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.apache.poi.hpsf.SummaryInformation;
import org.apache.poi.hssf.converter.AbstractExcelConverter;
import org.apache.poi.hssf.converter.ExcelToHtmlConverter;
import org.apache.poi.hssf.converter.ExcelToHtmlUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.hwpf.converter.HtmlDocumentFacade;
import org.apache.poi.ss.formula.eval.ErrorEval;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.POILogFactory;
import org.apache.poi.util.POILogger;
import org.springframework.util.StringUtils;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Text;

public class MyConverter extends AbstractExcelConverter {
	private static final POILogger logger = POILogFactory
			.getLogger(MyConverter.class);

	public static void main(String[] args) {
		if (args.length < 2) {
			System.err.println("Usage: ExcelToHtmlConverter <inputFile.xls> <saveTo.html>");
			return;
		}
		System.out.println("Converting " + args[0]);
		System.out.println("Saving output to " + args[1]);
		try {
			Document doc = process(new File(args[0]));

			FileWriter out = new FileWriter(args[1]);
			DOMSource domSource = new DOMSource(doc);
			StreamResult streamResult = new StreamResult(out);

			TransformerFactory tf = TransformerFactory.newInstance();
			Transformer serializer = tf.newTransformer();

			serializer.setOutputProperty("encoding", "UTF-8");
			serializer.setOutputProperty("indent", "no");
			serializer.setOutputProperty("method", "html");
			serializer.transform(domSource, streamResult);
			out.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public static Document process(File xlsFile) throws Exception {
		HSSFWorkbook workbook = ExcelToHtmlUtils.loadXls(xlsFile);
		ExcelToHtmlConverter excelToHtmlConverter = new ExcelToHtmlConverter(
				DocumentBuilderFactory.newInstance().newDocumentBuilder()
						.newDocument());

		excelToHtmlConverter.processWorkbook(workbook);
		return excelToHtmlConverter.getDocument();
	}

	private String cssClassContainerCell = null;
	private String cssClassContainerDiv = null;
	private String cssClassPrefixCell = "c";
	private String cssClassPrefixDiv = "d";
	private String cssClassPrefixRow = "r";
	private String cssClassPrefixTable = "t";
	private Map<Short, String> excelStyleToClass = new LinkedHashMap();
	private final HtmlDocumentFacade htmlDocumentFacade;
	private boolean useDivsToSpan = false;

	public MyConverter(Document doc) {
		this.htmlDocumentFacade = new HtmlDocumentFacade(doc);
	}

	public MyConverter(HtmlDocumentFacade htmlDocumentFacade) {
		this.htmlDocumentFacade = htmlDocumentFacade;
	}

	protected String buildStyle(HSSFWorkbook workbook, HSSFCellStyle cellStyle) {
		StringBuilder style = new StringBuilder();

		style.append("white-space:pre-wrap;");
		ExcelToHtmlUtils.appendAlign(style, cellStyle.getAlignment());
		if (cellStyle.getFillPattern() != 0) {
			if (cellStyle.getFillPattern() == 1) {
				HSSFColor foregroundColor = cellStyle
						.getFillForegroundColorColor();
				if (foregroundColor != null) {
					style.append("background-color:"
							+ ExcelToHtmlUtils.getColor(foregroundColor) + ";");
				}
			} else {
				HSSFColor backgroundColor = cellStyle
						.getFillBackgroundColorColor();
				if (backgroundColor != null) {
					style.append("background-color:"
							+ ExcelToHtmlUtils.getColor(backgroundColor) + ";");
				}
			}
		}
		buildStyle_border(workbook, style, "top", cellStyle.getBorderTop(),
				cellStyle.getTopBorderColor());

		buildStyle_border(workbook, style, "right", cellStyle.getBorderRight(),
				cellStyle.getRightBorderColor());

		buildStyle_border(workbook, style, "bottom",
				cellStyle.getBorderBottom(), cellStyle.getBottomBorderColor());

		buildStyle_border(workbook, style, "left", cellStyle.getBorderLeft(),
				cellStyle.getLeftBorderColor());

		HSSFFont font = cellStyle.getFont(workbook);
		buildStyle_font(workbook, style, font);

		return style.toString();
	}

	private void buildStyle_border(HSSFWorkbook workbook, StringBuilder style,
			String type, short xlsBorder, short borderColor) {
		if (xlsBorder == 0) {
			return;
		}
		StringBuilder borderStyle = new StringBuilder();
		borderStyle.append(ExcelToHtmlUtils.getBorderWidth(xlsBorder));
		borderStyle.append(' ');
		borderStyle.append(ExcelToHtmlUtils.getBorderStyle(xlsBorder));

		HSSFColor color = workbook.getCustomPalette().getColor(borderColor);
		if (color != null) {
			borderStyle.append(' ');
			borderStyle.append(ExcelToHtmlUtils.getColor(color));
		}
		style.append("border-" + type + ":" + borderStyle + ";");
	}

	void buildStyle_font(HSSFWorkbook workbook, StringBuilder style,
			HSSFFont font) {
		switch (font.getBoldweight()) {
		case 700:
			style.append("font-weight:bold;");
			break;
		}
		HSSFColor fontColor = workbook.getCustomPalette().getColor(
				font.getColor());
		if (fontColor != null) {
			style.append("color: " + ExcelToHtmlUtils.getColor(fontColor)
					+ "; ");
		}
		if (font.getFontHeightInPoints() != 0) {
			style.append("font-size:" + font.getFontHeightInPoints() + "pt;");
		}
		if (font.getItalic()) {
			style.append("font-style:italic;");
		}
	}

	public String getCssClassPrefixCell() {
		return this.cssClassPrefixCell;
	}

	public String getCssClassPrefixDiv() {
		return this.cssClassPrefixDiv;
	}

	public String getCssClassPrefixRow() {
		return this.cssClassPrefixRow;
	}

	public String getCssClassPrefixTable() {
		return this.cssClassPrefixTable;
	}

	public Document getDocument() {
		return this.htmlDocumentFacade.getDocument();
	}

	protected String getStyleClassName(HSSFWorkbook workbook,
			HSSFCellStyle cellStyle) {
		Short cellStyleKey = Short.valueOf(cellStyle.getIndex());

		String knownClass = (String) this.excelStyleToClass.get(cellStyleKey);
		if (knownClass != null) {
			return knownClass;
		}
		String cssStyle = buildStyle(workbook, cellStyle);
		String cssClass = this.htmlDocumentFacade.getOrCreateCssClass(
				this.cssClassPrefixCell, cssStyle);

		this.excelStyleToClass.put(cellStyleKey, cssClass);
		return cssClass;
	}

	public boolean isUseDivsToSpan() {
		return this.useDivsToSpan;
	}

	protected boolean processCell(HSSFCell cell, Element tableCellElement,
			int normalWidthPx, int maxSpannedWidthPx, float normalHeightPt) {
		HSSFCellStyle cellStyle = cell.getCellStyle();
		String value;
		switch (cell.getCellType()) {
		case 1:
			value = cell.getRichStringCellValue().getString();
			break;
		case 2:
			switch (cell.getCachedFormulaResultType()) {
			case 1:
				HSSFRichTextString str = cell.getRichStringCellValue();
				if ((str != null) && (str.length() > 0)) {
					value = str.toString();
				} else {
					value = "";
				}
				break;
			case 0:
				HSSFCellStyle style = cellStyle;
				if (style == null) {
					value = String.valueOf(cell.getNumericCellValue());
				} else {
					value = this._formatter.formatRawCellContents(
							cell.getNumericCellValue(), style.getDataFormat(),
							style.getDataFormatString());
				}
				break;
			case 4:
				value = String.valueOf(cell.getBooleanCellValue());
				break;
			case 5:
				value = ErrorEval.getText(cell.getErrorCellValue());
				break;
			case 2:
			case 3:
			default:
				logger.log(5, "Unexpected cell cachedFormulaResultType ("
						+ cell.getCachedFormulaResultType() + ")");

				value = "";
			}
			break;
		case 3:
			value = "";
			break;
		case 0:
			value = this._formatter.formatCellValue(cell);
			break;
		case 4:
			value = String.valueOf(cell.getBooleanCellValue());
			break;
		case 5:
			value = ErrorEval.getText(cell.getErrorCellValue());
			break;
		default:
			logger.log(5, "Unexpected cell type (" + cell.getCellType() + ")");

			return true;
		}

		boolean noText = StringUtils.isEmpty(value);
		boolean wrapInDivs = (!noText) && (isUseDivsToSpan())
				&& (!cellStyle.getWrapText());

		short cellStyleIndex = cellStyle.getIndex();
		if (cellStyleIndex != 0) {
			HSSFWorkbook workbook = cell.getRow().getSheet().getWorkbook();
			String mainCssClass = getStyleClassName(workbook, cellStyle);
			if (wrapInDivs) {
				tableCellElement.setAttribute("class", mainCssClass + " "
						+ this.cssClassContainerCell);
			} else {
				tableCellElement.setAttribute("class", mainCssClass);
			}
			if (noText) {
				value = "?";
			}
		}
		if ((isOutputLeadingSpacesAsNonBreaking()) && (value.startsWith(" "))) {
			StringBuilder builder = new StringBuilder();
			for (int c = 0; c < value.length(); c++) {
				if (value.charAt(c) != ' ') {
					break;
				}
				builder.append('?');
			}
			if (value.length() != builder.length()) {
				builder.append(value.substring(builder.length()));
			}
			value = builder.toString();
		}
		Text text = this.htmlDocumentFacade.createText(value);
		if (wrapInDivs) {
			Element outerDiv = this.htmlDocumentFacade.createBlock();
			outerDiv.setAttribute("class", this.cssClassContainerDiv);

			Element innerDiv = this.htmlDocumentFacade.createBlock();
			StringBuilder innerDivStyle = new StringBuilder();
			innerDivStyle.append("position:absolute;min-width:");
			innerDivStyle.append(normalWidthPx);
			innerDivStyle.append("px;");
			if (maxSpannedWidthPx != Integer.MAX_VALUE) {
				innerDivStyle.append("max-width:");
				innerDivStyle.append(maxSpannedWidthPx);
				innerDivStyle.append("px;");
			}
			innerDivStyle.append("overflow:hidden;max-height:");
			innerDivStyle.append(normalHeightPt);
			innerDivStyle.append("pt;white-space:nowrap;");
			ExcelToHtmlUtils.appendAlign(innerDivStyle,
					cellStyle.getAlignment());

			this.htmlDocumentFacade.addStyleClass(outerDiv,
					this.cssClassPrefixDiv, innerDivStyle.toString());

			innerDiv.appendChild(text);
			outerDiv.appendChild(innerDiv);
			tableCellElement.appendChild(outerDiv);
		} else {
			tableCellElement.appendChild(text);
		}
		return (StringUtils.isEmpty(value)) && (cellStyleIndex == 0);
	}

	protected void processColumnHeaders(HSSFSheet sheet, int maxSheetColumns,
			Element table) {
		Element tableHeader = this.htmlDocumentFacade.createTableHeader();
		table.appendChild(tableHeader);

		Element tr = this.htmlDocumentFacade.createTableRow();
		if (isOutputRowNumbers()) {
			tr.appendChild(this.htmlDocumentFacade.createTableHeaderCell());
		}
		for (int c = 0; c < maxSheetColumns; c++) {
			if ((isOutputHiddenColumns()) || (!sheet.isColumnHidden(c))) {
				Element th = this.htmlDocumentFacade.createTableHeaderCell();
				String text = getColumnName(c);
				th.appendChild(this.htmlDocumentFacade.createText(text));
				tr.appendChild(th);
			}
		}
		tableHeader.appendChild(tr);
	}

	protected void processColumnWidths(HSSFSheet sheet, int maxSheetColumns,
			Element table) {
		Element columnGroup = this.htmlDocumentFacade.createTableColumnGroup();
		if (isOutputRowNumbers()) {
			columnGroup
					.appendChild(this.htmlDocumentFacade.createTableColumn());
		}
		for (int c = 0; c < maxSheetColumns; c++) {
			if ((isOutputHiddenColumns()) || (!sheet.isColumnHidden(c))) {
				Element col = this.htmlDocumentFacade.createTableColumn();
				col.setAttribute("width",
						String.valueOf(getColumnWidth(sheet, c)));

				columnGroup.appendChild(col);
			}
		}
		table.appendChild(columnGroup);
	}

	protected void processDocumentInformation(
			SummaryInformation summaryInformation) {

		if (!StringUtils.isEmpty(summaryInformation.getTitle())) {
			this.htmlDocumentFacade.setTitle(summaryInformation.getTitle());
		}
		if (!StringUtils.isEmpty(summaryInformation.getAuthor())) {
			this.htmlDocumentFacade.addAuthor(summaryInformation.getAuthor());
		}
		if (!StringUtils.isEmpty(summaryInformation.getKeywords())) {
			this.htmlDocumentFacade.addKeywords(summaryInformation
					.getKeywords());
		}
		if (!StringUtils.isEmpty(summaryInformation.getComments())) {
			this.htmlDocumentFacade.addDescription(summaryInformation
					.getComments());
		}
	}

	protected int processRow(CellRangeAddress[][] mergedRanges, HSSFRow row,
			Element tableRowElement) {
		HSSFSheet sheet = row.getSheet();
		short maxColIx = row.getLastCellNum();
		if (maxColIx <= 0) {
			return 0;
		}
		List<Element> emptyCells = new ArrayList(maxColIx);
		if (isOutputRowNumbers()) {
			Element tableRowNumberCellElement = this.htmlDocumentFacade
					.createTableHeaderCell();

			processRowNumber(row, tableRowNumberCellElement);
			emptyCells.add(tableRowNumberCellElement);
		}
		int maxRenderedColumn = 0;
		for (int colIx = 0; colIx < maxColIx; colIx++) {
			if ((isOutputHiddenColumns()) || (!sheet.isColumnHidden(colIx))) {
				CellRangeAddress range = ExcelToHtmlUtils.getMergedRange(
						mergedRanges, row.getRowNum(), colIx);
				if ((range == null)
						|| ((range.getFirstColumn() == colIx) && (range
								.getFirstRow() == row.getRowNum()))) {
					HSSFCell cell = row.getCell(colIx);

					int divWidthPx = 0;
					if (isUseDivsToSpan()) {
						divWidthPx = getColumnWidth(sheet, colIx);

						boolean hasBreaks = false;
						for (int nextColumnIndex = colIx + 1; nextColumnIndex < maxColIx; nextColumnIndex++) {
							if ((isOutputHiddenColumns())
									|| (!sheet.isColumnHidden(nextColumnIndex))) {
								if ((row.getCell(nextColumnIndex) != null)
										&& (!isTextEmpty(row
												.getCell(nextColumnIndex)))) {
									hasBreaks = true;
									break;
								}
								divWidthPx += getColumnWidth(sheet,
										nextColumnIndex);
							}
						}
						if (!hasBreaks) {
							divWidthPx = Integer.MAX_VALUE;
						}
					}
					Element tableCellElement = this.htmlDocumentFacade
							.createTableCell();
					if (range != null) {
						if (range.getFirstColumn() != range.getLastColumn()) {
							tableCellElement.setAttribute(
									"colspan",
									String.valueOf(range.getLastColumn()
											- range.getFirstColumn() + 1));
						}
						if (range.getFirstRow() != range.getLastRow()) {
							tableCellElement.setAttribute(
									"rowspan",
									String.valueOf(range.getLastRow()
											- range.getFirstRow() + 1));
						}
					}
					boolean emptyCell;
					if (cell != null) {
						emptyCell = processCell(cell, tableCellElement,
								getColumnWidth(sheet, colIx), divWidthPx,
								row.getHeight() / 20.0F);
					} else {
						emptyCell = true;
					}
					if (emptyCell) {
						emptyCells.add(tableCellElement);
					} else {
						for (Element emptyCellElement : emptyCells) {
							tableRowElement.appendChild(emptyCellElement);
						}
						emptyCells.clear();

						tableRowElement.appendChild(tableCellElement);
						maxRenderedColumn = colIx;
					}
				}
			}
		}
		return maxRenderedColumn + 1;
	}

	protected void processRowNumber(HSSFRow row,
			Element tableRowNumberCellElement) {
		tableRowNumberCellElement.setAttribute("class", "rownumber");
		Text text = this.htmlDocumentFacade.createText(getRowName(row));
		tableRowNumberCellElement.appendChild(text);
	}

	protected void processSheet(HSSFSheet sheet) {
		processSheetHeader(this.htmlDocumentFacade.getBody(), sheet);

		int physicalNumberOfRows = sheet.getPhysicalNumberOfRows();
		if (physicalNumberOfRows <= 0) {
			return;
		}
		Element table = this.htmlDocumentFacade.createTable();
		this.htmlDocumentFacade.addStyleClass(table, this.cssClassPrefixTable,
				"border-collapse:collapse;border-spacing:0;");

		Element tableBody = this.htmlDocumentFacade.createTableBody();

		CellRangeAddress[][] mergedRanges = ExcelToHtmlUtils
				.buildMergedRangesMap(sheet);

		List<Element> emptyRowElements = new ArrayList(physicalNumberOfRows);

		int maxSheetColumns = 1;
		for (int r = sheet.getFirstRowNum(); r <= sheet.getLastRowNum(); r++) {
			HSSFRow row = sheet.getRow(r);
			if (row != null) {
				if ((isOutputHiddenRows()) || (!row.getZeroHeight())) {
					Element tableRowElement = this.htmlDocumentFacade
							.createTableRow();
					this.htmlDocumentFacade.addStyleClass(tableRowElement,
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

		this.htmlDocumentFacade.getBody().appendChild(table);
	}

	protected void processSheetHeader(Element htmlBody, HSSFSheet sheet) {
		Element h2 = this.htmlDocumentFacade.createHeader2();
		h2.appendChild(this.htmlDocumentFacade.createText(sheet.getSheetName()));
		htmlBody.appendChild(h2);
	}

	public void processWorkbook(HSSFWorkbook workbook) {
		SummaryInformation summaryInformation = workbook
				.getSummaryInformation();
		if (summaryInformation != null) {
			processDocumentInformation(summaryInformation);
		}
		if (isUseDivsToSpan()) {
			this.cssClassContainerCell = this.htmlDocumentFacade
					.getOrCreateCssClass(this.cssClassPrefixCell,
							"padding:0;margin:0;align:left;vertical-align:top;");

			this.cssClassContainerDiv = this.htmlDocumentFacade
					.getOrCreateCssClass(this.cssClassPrefixDiv,
							"position:relative;");
		}
		for (int s = 0; s < workbook.getNumberOfSheets(); s++) {
			HSSFSheet sheet = workbook.getSheetAt(s);
			processSheet(sheet);
		}
		this.htmlDocumentFacade.updateStylesheet();
	}

	public void setCssClassPrefixCell(String cssClassPrefixCell) {
		this.cssClassPrefixCell = cssClassPrefixCell;
	}

	public void setCssClassPrefixDiv(String cssClassPrefixDiv) {
		this.cssClassPrefixDiv = cssClassPrefixDiv;
	}

	public void setCssClassPrefixRow(String cssClassPrefixRow) {
		this.cssClassPrefixRow = cssClassPrefixRow;
	}

	public void setCssClassPrefixTable(String cssClassPrefixTable) {
		this.cssClassPrefixTable = cssClassPrefixTable;
	}

	public void setUseDivsToSpan(boolean useDivsToSpan) {
		this.useDivsToSpan = useDivsToSpan;
	}
}
