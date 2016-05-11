package com.vjd.dtutil;

import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.InputStream;

import jxl.Sheet;
import jxl.Workbook;
/**
 * 中国各区县数据初始化程序
 * 要保证excel文档的格式规范
 * 增加补齐和空行功能
 * @author 陈清水
 *
 */
public class GenExcelSql {
	
	private static String readFile="d:\\中国各个市县名称汇总.xls"; //要读取的县区文件（Excel文档）
	private static String saveSqlf="d:\\mquxian.sql"; //要保存的县区分类SQL（文本文件）
	private static int maxLen=25;//名称最大长度

	public static void main(String[] args) {
		try {
			Workbook rwb = null;
			// 创建输入流
			InputStream stream = new FileInputStream(readFile);
			// 获取Excel文件对象
			rwb = Workbook.getWorkbook(stream);
			// 获取文件的指定工作表 默认的第一个
			Sheet sheet = rwb.getSheet(0);
			int rowct = sheet.getRows();
			String province=null;//省名称
			String city = null;//城市
			String area = null;//地区
			String code = null;//编码
			StringBuilder sql = new StringBuilder();
			int j=1;//初始化code序号
			for (int i = 1; i < rowct; i++) {
				province=sheet.getCell(3, i).getContents().trim();
				city=sheet.getCell(2, i).getContents().trim();
				area=sheet.getCell(1, i).getContents().trim();
				if(i>1){
					if(city.equals(sheet.getCell(2, i-1).getContents()))//当前城市与上一行城市相等
						j=j+1;//编号+1
					else //不相等从新开始编号
						j=1;
				}
				if(j<10) //编码补0
					code="0"+j;
				else
					code=String.valueOf(j);
				int len = maxLen - 2*area.length();
				String buqi = "";
				for(int k=1;k<len;k++)
					buqi=buqi+" ";
				if(j==1&&i>1)
					sql.append("\r\n");//从新编号时补空行
				if(!city.equals("省直辖"))//省直辖有重复数据
					sql.append("insert into mquxian(code,name,city_id) values('"+code+"','"+area+"'"+buqi+", (select id from mshi where name='"+city+"' limit 1));\r\n");
				else
					sql.append("insert into mquxian(code,name,city_id) values('"+code+"','"+area+"'"+buqi+", (select id from mshi where name='"+city+"' and sh_id = (select id from msheng where name='"+province+"') limit 1));\r\n");
			}
			FileWriter fw = new FileWriter(saveSqlf);//保存的sql文件
			fw.write(sql.toString());//写入数据
			fw.close();//关闭
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

}
