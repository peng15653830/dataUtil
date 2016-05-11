package com.vjd.dtutil;

import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.InputStream;

import jxl.Sheet;
import jxl.Workbook;

public class GenShangpSql {

	private static String readFile="d:\\商品数据录入模板.xls"; 
	private static String saveSqlf="d:\\shangpinshebeijishucanshu.sql"; 
		
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
			String changjia=null;//厂家名称
			String fl1=null;//产品大类
			String fl2 = null;//产品中类
			String fl3 = null;//产品小类
			
			String shangpincode=null;//商品编码
			String shangpinname=null;//商品名称
			String shangpinpinpai=null;//商品品牌
			String baozhuangqingdan=null;//包装清单kbtj
			String shouhoubaozhang=null;//售后保障shhbz
			
			String huoqi=null;//货期(天)
			String shengchanhuoqi=null;//生产货期(天)
			
			String xinghao=null;//型号
			String gonghuojia=null;//供货价
			String zhongliang=null;//重量
			String chang=null;//长
			String kuan=null;//宽
			String gao=null;//高
			String tupian=null;//图片名
			
			StringBuilder sql = new StringBuilder();
			for (int i = 1; i < rowct; i++) {
				shangpinname=sheet.getCell(5, i).getContents().trim();
				
				if(shangpinname != null && !shangpinname.equals("")){
					fl3=sheet.getCell(3, i).getContents().trim();
					fl2=sheet.getCell(2, i).getContents().trim();
					fl1=sheet.getCell(1, i).getContents().trim();
					changjia=sheet.getCell(0, i).getContents().trim();
					shangpincode=sheet.getCell(4, i).getContents().trim();
					
					shangpinpinpai=sheet.getCell(6, i).getContents().trim();
					baozhuangqingdan=sheet.getCell(7, i).getContents().trim();
					shouhoubaozhang=sheet.getCell(8, i).getContents().trim();
					huoqi=sheet.getCell(9, i).getContents().trim();
					shengchanhuoqi=sheet.getCell(10, i).getContents().trim();
					xinghao=sheet.getCell(11, i).getContents().trim();
					gonghuojia=sheet.getCell(12, i).getContents().trim();
					zhongliang=sheet.getCell(13, i).getContents().trim();
					chang=sheet.getCell(14, i).getContents().trim();
					kuan=sheet.getCell(15, i).getContents().trim();
					gao=sheet.getCell(16, i).getContents().trim();
					tupian=sheet.getCell(17, i).getContents().trim();
					
					sql.append("insert into mshangp(shpfl1_id,shpfl2_id,shpfl3_id,chj_id,code,name,kbtj,shhbz,huoqi1,huoqi2,flag,org_id,op_id,op_dpt,op_ip) values((select m1.id from mshpfl1 m1 where m1.name='"+fl1+"' and m1.flag = 1),(select m2.id from mshpfl2 m2 where m2.name ='"+fl2+"' and m2.shpfl1_id=(select m1.id from mshpfl1 m1 where m1.name='"+fl1+"' and m1.flag = 1)),(select m3.id from mshpfl3 m3 where m3.name = '"+fl3+"' and m3.shpfl2_id = (select m2.id from mshpfl2 m2 where m2.name ='"+fl2+"' and m2.shpfl1_id=(select m1.id from mshpfl1 m1 where m1.name='"+fl1+"' and m1.flag = 1))),(SELECT mcj.id from mchangj mcj where mcj.name = '"+changjia+"' and mcj.flag = 1),'"+shangpincode+"','"+shangpinname+"','"+baozhuangqingdan+"','"+shouhoubaozhang+"',0,0,2,0,0,0,'0:0:0:0:0:0:0:1');\r\n");
					System.out.println("厂家:"+changjia+" 商品编号:"+shangpincode+" 商品名称:"+shangpinname+" 商品品牌:"+shangpinpinpai+" 捆包条件:"+baozhuangqingdan+" 售后保障:"+shouhoubaozhang+" 货期(天):"+huoqi+" 生产货期(天):"+shengchanhuoqi+" 型号:"+xinghao+" 供货价:"+gonghuojia+" 重量:"+zhongliang+" 长:"+chang+" 宽:"+kuan+" 高:"+gao+" 产品大类:"+fl1+" 产品中类:"+fl2+" 产品小类:"+fl3);
					//sql.append("insert into mquxian(code,name,city_id) values('"+code+"','"+area+"'"+buqi+", (select id from mshi where name='"+city+"' and sh_id = (select id from msheng where name='"+province+"') limit 1));\r\n");
					
					FileWriter fw = new FileWriter(saveSqlf);//保存的sql文件
					fw.write(sql.toString());//写入数据
					fw.close();//关闭
				}
				
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

}
