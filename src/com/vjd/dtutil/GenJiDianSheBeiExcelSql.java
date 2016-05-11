package com.vjd.dtutil;

import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.InputStream;

import jxl.Sheet;
import jxl.Workbook;

public class GenJiDianSheBeiExcelSql {

	private static String readFile="d:\\机电设备技术参数1.xls"; //要读取的县区文件（Excel文档）
	//private static String readFile="d:\\技术参数录入（电器电缆）.xls"; //要读取的县区文件（Excel文档）
	private static String saveSqlf="d:\\jidianshebei.sql"; //要保存的县区分类SQL（文本文件）
	//private static String saveSqlf="d:\\jidianshebei111.sql"; //要保存的县区分类SQL（文本文件）
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
			String fl1=null;//产品大类
			String fl2 = null;//产品中类
			String fl3 = null;//产品小类
			String ltext = null;//技术参数名称(内容)
			String lkind = null;//类别(0 标题 1标签)
			String bitian = "0";//必填(1 必填 0 不必填)
			String jdisp = null;//商品介绍(0 无介绍 1有介绍)
			String selitem = null;//可筛选(0 无筛选 1有筛选)
			String seltext = null;//筛选值
			String deftext = null;//初始值(默认值)
			StringBuilder sql = new StringBuilder();
			for (int i = 1; i < rowct; i++) {
				fl3=sheet.getCell(2, i).getContents().trim();
				fl2=sheet.getCell(1, i).getContents().trim();
				fl1=sheet.getCell(0, i).getContents().trim();
				
				if(fl1 != null && !"".equals(fl1)){
					ltext=sheet.getCell(3, i).getContents().trim();
					lkind="1";
					if(sheet.getCell(4, i).getContents() != null){
						jdisp=sheet.getCell(4, i).getContents().trim();
					}
					
					//selitem=sheet.getCell(3, i).getContents().trim();					
					if(sheet.getCell(5, i).getContents() != null){
						deftext=sheet.getCell(5, i).getContents().trim();
					}
					
					//筛选值
					if(sheet.getCell(6, i).getContents() != null){
						seltext=sheet.getCell(6, i).getContents().trim();
						for(int j = 1; j <= 10; j++){
							//String aa = sheet.getCell(6+j, i).getContents().trim();
							//System.out.println("qqqq: "+aa);
							System.out.println("qqqq: "+j);
						}
					}
								
					//商品介绍
					if(jdisp.equals("是")){
						jdisp = "1";
					}else{
						jdisp = "0";
					}
					
					//可筛选
					if(seltext != null){
						selitem = "1";
					}else{
						selitem = "0";
					}
									
					sql.append("insert into mshpggmb(shpfl3_id,ltext,lkind,bitian,jdisp,selitem,deftext,op_ip,seltext) values((select m3.id from mshpfl3 m3 where m3.name = '"+fl3+"' and m3.shpfl2_id = (select m2.id from mshpfl2 m2 where m2.name ='"+fl2+"' and m2.shpfl1_id=(select m1.id from mshpfl1 m1 where m1.name='"+fl1+"'))),'"+ltext+"',"+lkind+","+bitian+","+jdisp+","+selitem+",'"+deftext+"','0:0:0:0:0:0:0:1','"+seltext+"');\r\n");
					System.out.println("产品大类:"+fl1+" 产品中类:"+fl2+" 产品小类:"+fl3+" 技术参数名称:"+ltext+" 商品介绍:"+jdisp+" 默认值:"+deftext+" 筛选值:"+seltext);
					//sql.append("insert into mquxian(code,name,city_id) values('"+code+"','"+area+"'"+buqi+", (select id from mshi where name='"+city+"' and sh_id = (select id from msheng where name='"+province+"') limit 1));\r\n");
				
				}
				
				}
			FileWriter fw = new FileWriter(saveSqlf);//保存的sql文件
			fw.write(sql.toString());//写入数据
			fw.close();//关闭
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

}
