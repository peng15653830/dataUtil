package com.vjd.dtutil;

import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.InputStream;

import jxl.Sheet;
import jxl.Workbook;

public class GenShangpSql {

	private static String readFile="d:\\��Ʒ����¼��ģ��.xls"; 
	private static String saveSqlf="d:\\shangpinshebeijishucanshu.sql"; 
		
	public static void main(String[] args) {

		try {
			Workbook rwb = null;
			// ����������
			InputStream stream = new FileInputStream(readFile);
			// ��ȡExcel�ļ�����
			rwb = Workbook.getWorkbook(stream);
			// ��ȡ�ļ���ָ�������� Ĭ�ϵĵ�һ��
			Sheet sheet = rwb.getSheet(0);
			int rowct = sheet.getRows();
			String changjia=null;//��������
			String fl1=null;//��Ʒ����
			String fl2 = null;//��Ʒ����
			String fl3 = null;//��ƷС��
			
			String shangpincode=null;//��Ʒ����
			String shangpinname=null;//��Ʒ����
			String shangpinpinpai=null;//��ƷƷ��
			String baozhuangqingdan=null;//��װ�嵥kbtj
			String shouhoubaozhang=null;//�ۺ���shhbz
			
			String huoqi=null;//����(��)
			String shengchanhuoqi=null;//��������(��)
			
			String xinghao=null;//�ͺ�
			String gonghuojia=null;//������
			String zhongliang=null;//����
			String chang=null;//��
			String kuan=null;//��
			String gao=null;//��
			String tupian=null;//ͼƬ��
			
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
					System.out.println("����:"+changjia+" ��Ʒ���:"+shangpincode+" ��Ʒ����:"+shangpinname+" ��ƷƷ��:"+shangpinpinpai+" ��������:"+baozhuangqingdan+" �ۺ���:"+shouhoubaozhang+" ����(��):"+huoqi+" ��������(��):"+shengchanhuoqi+" �ͺ�:"+xinghao+" ������:"+gonghuojia+" ����:"+zhongliang+" ��:"+chang+" ��:"+kuan+" ��:"+gao+" ��Ʒ����:"+fl1+" ��Ʒ����:"+fl2+" ��ƷС��:"+fl3);
					//sql.append("insert into mquxian(code,name,city_id) values('"+code+"','"+area+"'"+buqi+", (select id from mshi where name='"+city+"' and sh_id = (select id from msheng where name='"+province+"') limit 1));\r\n");
					
					FileWriter fw = new FileWriter(saveSqlf);//�����sql�ļ�
					fw.write(sql.toString());//д������
					fw.close();//�ر�
				}
				
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

}
