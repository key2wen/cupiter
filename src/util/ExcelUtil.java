package com.to8to.weixin.util;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.Properties;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.thrift.TBase;
import org.apache.thrift.TFieldIdEnum;
import org.apache.thrift.meta_data.FieldMetaData;

import com.to8to.commons.utils.Config;
import com.to8to.weixin.thrift.TActivity;
import com.to8to.weixin.thrift.TMsg;
import com.to8to.weixin.thrift.TUser;

public class ExcelUtil {

	static Properties properties = null;
	static Config config = null;
	private int rowNum = 0;
	private HSSFWorkbook workbook = null;

	public ExcelUtil() {
		super();
		initProperties();
		workbook = new HSSFWorkbook();
	}

	public static void initProperties() {
		if (properties == null) {
			properties = new Properties();
			try {
				properties.load(new InputStreamReader(ExcelUtil.class
						.getClassLoader().getResourceAsStream(
								"excel.properties"), "UTF-8"));
			} catch (Exception e) {
				e.printStackTrace();
			}
		}
	}

	public static void initConfig() {
		if (config == null) {
			config = new Config("excel.properties");
		}

	}

	public void exportMsg(List<TMsg> msgs) {
		String[] header = properties.getProperty("excel.msg.header").split(",");
		String excelName = properties.getProperty("excel.msg.name");
		String sheetName = properties.getProperty("excel.msg.title");

		String pattern = properties.getProperty("excel.date.pattern");
		String exportPath = properties.getProperty("excel.exportPath");

		HSSFSheet sheet = workbook.createSheet(sheetName);
		sheet.setDefaultColumnWidth((short) 15);
		HSSFCellStyle style = getHeaderStyle();
		this.createHeader(sheet, style, header);

		handleData(msgs, sheet, pattern, exportPath, excelName);
	}

	public void exportActivity(List<TActivity> activitys) {
		String[] header = properties.getProperty("excel.activity.header")
				.split(",");
		String excelName = properties.getProperty("excel.activity.name");
		String sheetName = properties.getProperty("excel.activity.title");

		String pattern = properties.getProperty("excel.date.pattern");
		String exportPath = properties.getProperty("excel.exportPath");

		HSSFSheet sheet = workbook.createSheet(sheetName);
		sheet.setDefaultColumnWidth((short) 15);
		HSSFCellStyle style = getHeaderStyle();
		this.createHeader(sheet, style, header);

		handleData(activitys, sheet, pattern, exportPath, excelName);
	}

	public void exportActivityUser(List<TUser> users, String activity_id,
			String activity_name) {
		String[] header = properties.getProperty("excel.activityUser.header")
				.split(",");
		String excelName = properties.getProperty("excel.activityUser.name");
		String sheetName = properties.getProperty("excel.activityUser.title");
		String activityId = properties
				.getProperty("excel.activityUser.activityId");
		String activityName = properties
				.getProperty("excel.activityUser.activityName");

		String[] headerTile = { activityId, activityName };
		String[] headerContent = { activity_id, activity_name };
		String pattern = properties.getProperty("excel.date.pattern");
		String exportPath = properties.getProperty("excel.exportPath");

		HSSFSheet sheet = workbook.createSheet(sheetName);
		sheet.setDefaultColumnWidth((short) 15);
		HSSFCellStyle style = getHeaderStyle();
		HSSFCellStyle contentStyle = getContentStyle(0);

		this.createHeader(sheet, style, headerTile);
		this.createHeader(sheet, contentStyle, headerContent);
		this.createHeader(sheet, style, header);

		handleData(users, sheet, pattern, exportPath, excelName);
	}

	public HSSFRow createRow(HSSFSheet sheet) {
		HSSFRow row = sheet.createRow(rowNum);
		rowNum++;
		return row;
	}

	private void createHeader(HSSFSheet sheet, HSSFCellStyle style,
			String[] header) {
		if (header == null)
			return;
		HSSFRow headerRow = this.createRow(sheet);
		short i = 0;
		for (String head : header) {
			HSSFCell cell = headerRow.createCell(i++);
			cell.setCellStyle(style);
			HSSFRichTextString text = new HSSFRichTextString(head);
			cell.setCellValue(text);
		}
	}

	public <F extends TFieldIdEnum, T extends TBase> void handleData(
			List dataList, HSSFSheet sheet, String pattern, String exportPath,
			String excelName) {
		OutputStream out = null;
		try {
			out = new FileOutputStream(exportPath + File.separatorChar
					+ excelName);
			if (dataList == null || dataList.size() <= 0) {
				workbook.write(out);
				return;
			}
			// 声明一个画图的顶级管理器
			HSSFPatriarch patriarch = sheet.createDrawingPatriarch();
			Object object = dataList.get(0);
			if (object instanceof TBase) {
				thriftDataHandle(dataList, sheet, pattern, patriarch, object);
			} else {
				if (object.getClass().isArray()) {
					imageHandle(dataList, sheet, pattern, patriarch, object);
				} else {
					genericObject(dataList, sheet, pattern, patriarch, object);
				}
			}
			workbook.write(out);
			ByteArrayOutputStream os = new ByteArrayOutputStream(); 
			os.toByteArray();
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			try {
				out.flush();
				out.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}

	public void imageHandle(List dataList, HSSFSheet sheet, String pattern,
			HSSFPatriarch patriarch, Object object) {
		HSSFRow row = null;
		if (object instanceof byte[]) {
			short col_num = 0;
			row = createRow(sheet);
			for (Object data : dataList) {
				byte[] image = (byte[]) data;
				HSSFCell cell = row.createCell(col_num);
				setValueToCell(sheet, row, cell, image, pattern, col_num,
						patriarch);
				col_num++;
			}
		}
	}

	// 普通java对象通过反射拿数据
	public void genericObject(List dataList, HSSFSheet sheet, String pattern,
			HSSFPatriarch patriarch, Object object) {
		HSSFRow row = null;
		for (Object o : dataList) {
			row = this.createRow(sheet);
			Field[] fields = o.getClass().getDeclaredFields();
			short col_num = 0;
			for (Field field : fields) {
				String fieldName = field.getName();
				String getMethodName = "get" + fieldName.substring(0, 1).toUpperCase()
						+ fieldName.substring(1);
				Class cls = o.getClass();
				Method method;
				try {
					method = cls.getMethod(getMethodName, new Class[] {});
					Object value = method.invoke(o, new Object[] {});
					HSSFCell cell = row.createCell(col_num);
					setValueToCell(sheet, row, cell, value, pattern, col_num,
							patriarch);
					col_num++;
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		}
	}

	public <T extends TBase, F extends TFieldIdEnum> void thriftDataHandle(
			List dataList, HSSFSheet sheet, String pattern,
			HSSFPatriarch patriarch, Object object)
			throws NoSuchFieldException, IllegalAccessException {
		HSSFRow row = null;
		// 取得metaDataMap属性
		Field metaDataMapField = object.getClass().getField("metaDataMap");
		// metaDataMap的类型一定是：Map<F,FieldMetaData>，所以这里的强转是安全的
		Map<F, FieldMetaData> metaDataMap = (Map<F, FieldMetaData>) metaDataMapField
				.get(object);
		for (Object data : dataList) {
			T thriftObject_data = (T) data;
			short col_num = 0;
			row = createRow(sheet);
			for (F field : metaDataMap.keySet()) {
				Object fieldValue = thriftObject_data.getFieldValue(field);
				HSSFCell cell = row.createCell(col_num);
				setValueToCell(sheet, row, cell, fieldValue, pattern, col_num,
						patriarch);
				col_num++;
			}
		}
	}

	private void setValueToCell(HSSFSheet sheet, HSSFRow row, HSSFCell cell,
			Object value, String pattern, short col_num, HSSFPatriarch patriarch) {
		if (value == null)
			value = "";
		// 判断值的类型后进行强制类型转换
		String textValue = null;
		if (value instanceof Boolean) {
			boolean bValue = (Boolean) value;
			textValue = "true";
			if (!bValue) {
				textValue = "false";
			}
		} else if (value instanceof Date) {
			Date date = (Date) value;
			SimpleDateFormat sdf = new SimpleDateFormat(pattern);
			textValue = sdf.format(date);
		} else if (value instanceof byte[]) {
			// 有图片时，设置行高为60px;
			row.setHeightInPoints(60);
			// 设置图片所在列宽度为80px
			sheet.setColumnWidth(col_num, (short) (35.7 * 80));
			// sheet.autoSizeColumn(col_num);
			byte[] imageValue = (byte[]) value;
			HSSFClientAnchor anchor = new HSSFClientAnchor(0, 0, 1023, 255,
					col_num, rowNum - 1, col_num, rowNum - 1);
			anchor.setAnchorType(2);
			patriarch.createPicture(anchor, workbook.addPicture(imageValue,
					HSSFWorkbook.PICTURE_TYPE_JPEG));
		} else {
			// 其它数据类型都当作字符串简单处理
			textValue = value.toString();
		}
		// 如果不是图片数据，就利用正则表达式判断textValue是否全部由数字组成
		if (textValue != null) {
			Pattern p = Pattern.compile("^//d+(//.//d+)?$");
			Matcher matcher = p.matcher(textValue);
			if (matcher.matches()) {
				// 是数字当作double处理
				cell.setCellValue(Double.parseDouble(textValue));
			} else {
				HSSFRichTextString richString = new HSSFRichTextString(
						textValue);
				HSSFFont fontContent = workbook.createFont();
				fontContent.setColor(HSSFColor.BLUE.index);
				richString.applyFont(fontContent);
				cell.setCellValue(richString);
			}
			HSSFCellStyle contentStyle = getContentStyle(rowNum - 1);
			cell.setCellStyle(contentStyle);
		}
	}

	private HSSFCellStyle getContentStyle(int row) {
		HSSFCellStyle contentStyle = workbook.createCellStyle();
		if (row % 2 == 0) {
			contentStyle.setFillForegroundColor(HSSFColor.LIGHT_YELLOW.index);
		}
		contentStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		contentStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		contentStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		contentStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
		contentStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
		contentStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		contentStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
		HSSFFont conentFont = workbook.createFont();
		conentFont.setBoldweight(HSSFFont.BOLDWEIGHT_NORMAL);
		contentStyle.setFont(conentFont);
		return contentStyle;
	}

	private HSSFCellStyle getHeaderStyle() {
		HSSFCellStyle headerStyle = workbook.createCellStyle();
		headerStyle.setFillForegroundColor(HSSFColor.SKY_BLUE.index);
		headerStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		headerStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		headerStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		headerStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
		headerStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
		headerStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		HSSFFont font = workbook.createFont();
		font.setColor(HSSFColor.VIOLET.index);
		font.setFontHeightInPoints((short) 15);
		font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
		headerStyle.setFont(font);

		return headerStyle;
	}

	public void exportImage(List<byte[]> bytes) {
		String[] header = properties.getProperty("excel.msg.header").split(",");
		String excelName = properties.getProperty("excel.msg.name");
		String sheetName = properties.getProperty("excel.msg.title");

		String pattern = properties.getProperty("excel.date.pattern");
		String exportPath = properties.getProperty("excel.exportPath");

		HSSFSheet sheet = workbook.createSheet(sheetName);
		sheet.setDefaultColumnWidth((short) 15);
		HSSFCellStyle style = getHeaderStyle();
		this.createHeader(sheet, style, header);

		handleData(bytes, sheet, pattern, exportPath, excelName);
	}

	public void exportGenericBean(List genericBeans){
		String[] header = properties.getProperty("excel.msg.header").split(",");
		String excelName = properties.getProperty("excel.msg.name");
		String sheetName = properties.getProperty("excel.msg.title");

		String pattern = properties.getProperty("excel.date.pattern");
		String exportPath = properties.getProperty("excel.exportPath");

		HSSFSheet sheet = workbook.createSheet(sheetName);
		sheet.setDefaultColumnWidth((short) 15);
		HSSFCellStyle style = getHeaderStyle();
		this.createHeader(sheet, style, header);

		handleData(genericBeans, sheet, pattern, exportPath, excelName);
	}
	
}
