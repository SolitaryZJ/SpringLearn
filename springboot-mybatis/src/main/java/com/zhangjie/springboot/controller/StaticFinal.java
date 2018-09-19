package com.zhangjie.springboot.controller;

/**
 * 估值表导入常量信息
 *
 * @version 1.0
 * @since JDK1.7
 * @author leipan
 * @company 上海朝阳永续信息技术有限公司
 * @Date 2017/12/7 15:55
 */
public class StaticFinal {

	public static final String[] DATE_FORMAT_LIST = {"yyyy-MM-dd", "yyyy/MM/dd", "yyyyMMdd", "yyyy年MM月dd日"};
	public static final String[] CODE_COLUMN_NAME = {"科目代码"};
	public static final String[] NAME_COLUMN_NAME = {"科目名称"};
	public static final String[] NUM_COLUMN_NAME = {"数量"};
	public static final String[] UNIT_COST_COLUMN_NAME = {"单位成本"};
	public static final String[] COST_COLUMN_NAME = {"成本", "成本-本币", "成本/本币"};
	public static final String[] COST_POR_COLUMN_NAME = {"成本占净值", "成本占比", "占资产净值"};
	public static final String[] PRICE_COLUMN_NAME = {"市价", "行情", "行情收市价"};
	public static final String[] MARKET_VALUE_COLUMN_NAME = {"市值", "市值-本币", "市值/本币"};
	public static final String[] MARKET_VALUE_PRO_COLUMN_NAME = {"市值占比", "市值占净值", "占资产净值"};
	public static final String[] VALUATION_INCR_COLUMN_NAME = {"估值增值", "估值增值-本币"};
	public static final String[] STOP_INFO_COLUMN_NAME = {"停牌信息", "停牌标准"};

	public static final String[] NAV_COLUMN_NAME = {"基金单位净值", "资产单位净值", "今日单位净值", "基金资产单位净值", "单位净值"};
	public static final String[] ADDED_NAV_COLUMN_NAME = {"累计单位净值", "累计基金净值"};
	public static final String[] TOTAL_NUMBER_NAME = {"实收资本"};
	public static final String[] PAID_IN_CAPITAL_COLUMN_NAME = {"资产类合计", "资产合计"};
	public static final String[] DEBT_COLUMN_NAME = {"负债类合计", "负债合计"};
	public static final String[] ASSET_NAV_COLUMN_NAME = {"基金资产净值", "资产净值", "信托资产净值", "集合计划资产净值"};

	public static final Long Max_File_size = 15*1024*1024L; //压缩包最大值不超过 15M
	public static final Long Max_Single_File_size = 500*1024L; //单个文件最大值 不超过 500k
	public static final Long MAX_LINE = 5000L; //文件最大不能超过5000行
	public static final String[] FILE_EXCEL_SUFFIX = {"xls","XLS", "xlsx", "XLSX", "xlsm", "XLSM", "xltx", "XLTX", "xltm", "XLTM", "xlsb", "XLSB",  "xlam", "XLAM"};

	public static final String FILE_PATH = "tmp/data"; //估值上传的缓存路径
}
