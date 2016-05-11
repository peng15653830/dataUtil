package com.vjd.dtutil;

import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.InputStream;

import jxl.Sheet;
import jxl.Workbook;

public class GenJiDianSheBeiExcelSql {

	private static String readFile="d:\\�����豸��������1.xls"; //Ҫ��ȡ�������ļ���Excel�ĵ���
	//private static String readFile="d:\\��������¼�루�������£�.xls"; //Ҫ��ȡ�������ļ���Excel�ĵ���
	private static String saveSqlf="d:\\jidianshebei.sql"; //Ҫ�������������SQL���ı��ļ���
	//private static String saveSqlf="d:\\jidianshebei111.sql"; //Ҫ�������������SQL���ı��ļ���
	private static int maxLen=25;//������󳤶�

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
			String fl1=null;//��Ʒ����
			String fl2 = null;//��Ʒ����
			String fl3 = null;//��ƷС��
			String ltext = null;//������������(����)
			String lkind = null;//���(0 ���� 1��ǩ)
			String bitian = "0";//����(1 ���� 0 ������)
			String jdisp = null;//��Ʒ����(0 �޽��� 1�н���)
			String selitem = null;//��ɸѡ(0 ��ɸѡ 1��ɸѡ)
			String seltext = null;//ɸѡֵ
			String deftext = null;//��ʼֵ(Ĭ��ֵ)
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
					
					//ɸѡֵ
					if(sheet.getCell(6, i).getContents() != null){
						seltext=sheet.getCell(6, i).getContents().trim();
						for(int j = 1; j <= 10; j++){
							//String aa = sheet.getCell(6+j, i).getContents().trim();
							//System.out.println("qqqq: "+aa);
							System.out.println("qqqq: "+j);
						}
					}
								
					//��Ʒ����
					if(jdisp.equals("��")){
						jdisp = "1";
					}else{
						jdisp = "0";
					}
					
					//��ɸѡ
					if(seltext != null){
						selitem = "1";
					}else{
						selitem = "0";
					}
									
					sql.append("insert into mshpggmb(shpfl3_id,ltext,lkind,bitian,jdisp,selitem,deftext,op_ip,seltext) values((select m3.id from mshpfl3 m3 where m3.name = '"+fl3+"' and m3.shpfl2_id = (select m2.id from mshpfl2 m2 where m2.name ='"+fl2+"' and m2.shpfl1_id=(select m1.id from mshpfl1 m1 where m1.name='"+fl1+"'))),'"+ltext+"',"+lkind+","+bitian+","+jdisp+","+selitem+",'"+deftext+"','0:0:0:0:0:0:0:1','"+seltext+"');\r\n");
					System.out.println("��Ʒ����:"+fl1+" ��Ʒ����:"+fl2+" ��ƷС��:"+fl3+" ������������:"+ltext+" ��Ʒ����:"+jdisp+" Ĭ��ֵ:"+deftext+" ɸѡֵ:"+seltext);
					//sql.append("insert into mquxian(code,name,city_id) values('"+code+"','"+area+"'"+buqi+", (select id from mshi where name='"+city+"' and sh_id = (select id from msheng where name='"+province+"') limit 1));\r\n");
				
				}
				
				}
			FileWriter fw = new FileWriter(saveSqlf);//�����sql�ļ�
			fw.write(sql.toString());//д������
			fw.close();//�ر�
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

}
