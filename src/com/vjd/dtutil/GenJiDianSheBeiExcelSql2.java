package com.vjd.dtutil;

import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.InputStream;

import jxl.Sheet;
import jxl.Workbook;

public class GenJiDianSheBeiExcelSql2 {

	//private static String readFile="d:\\�����豸��������1.xls"; //Ҫ��ȡ�������ļ���Excel�ĵ���
	//private static String readFile="d:\\��������¼�루�������£�.xls"; //Ҫ��ȡ�������ļ���Excel�ĵ���
	//private static String readFile="d:\\��������¼�루�����豸��.xls"; //Ҫ��ȡ�������ļ���Excel�ĵ���
	//private static String readFile="d:\\��������¼�루�����豸2��.xls"; //Ҫ��ȡ�������ļ���Excel�ĵ���
	private static String readFile="d:\\��������¼�루�ͱ�������.xls"; //Ҫ��ȡ�������ļ���Excel�ĵ�
	//private static String readFile="d:\\��������¼�루�����Ǳ�.xls"; //Ҫ��ȡ�������ļ���Excel�ĵ�
	//private static String readFile="d:\\��������¼�루ˮů���ģ�.xls"; //Ҫ��ȡ�������ļ���Excel�ĵ�
	//private static String readFile="d:\\��������¼�루ˮů����2��.xls"; //Ҫ��ȡ�������ļ���Excel�ĵ�
	//private static String readFile="d:\\��������¼�루ˮů����3��.xls"; //Ҫ��ȡ�������ļ���Excel�ĵ�
	
	//private static String saveSqlf="d:\\jidianshebei.sql"; //Ҫ�������������SQL���ı��ļ���
	//private static String saveSqlf="d:\\��������¼�루�������£�22.sql"; //Ҫ�������������SQL���ı��ļ���
	//private static String saveSqlf="d:\\��������¼�루�����豸��.sql"; //Ҫ�������������SQL���ı��ļ�
	//private static String saveSqlf="d:\\��������¼�루�����豸2��.sql"; //Ҫ�������������SQL���ı��ļ���
	private static String saveSqlf="d:\\��������¼�루�ͱ�������.sql"; //Ҫ�������������SQL���ı��ļ���
	//private static String saveSqlf="d:\\��������¼�루�����Ǳ�.sql"; //Ҫ�������������SQL���ı��ļ���
	//private static String saveSqlf="d:\\��������¼�루ˮů���ģ�.sql"; //Ҫ�������������SQL���ı��ļ���
	//private static String saveSqlf="d:\\��������¼�루ˮů����2��.sql"; //Ҫ�������������SQL���ı��ļ���
	//private static String saveSqlf="d:\\��������¼�루ˮů����3��.sql"; //Ҫ�������������SQL���ı��ļ���
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
			String fl3value = "";//��ƷС��value
			String ltext = null;//������������(����)
			String lkind = null;//���(0 ���� 1��ǩ)
			String bitian = "0";//����(1 ���� 0 ������)
			String jdisp = null;//��Ʒ����(0 �޽��� 1�н���)
			String selitem = null;//��ɸѡ(0 ��ɸѡ 1��ɸѡ)
			String seltext = null;//ɸѡֵ
			String deftext = null;//��ʼֵ(Ĭ��ֵ)
			StringBuilder sql = new StringBuilder();
			StringBuilder sql2 = new StringBuilder();
			for (int i = 1; i < rowct; i++) {
				fl3=sheet.getCell(2, i).getContents().trim();
				if(i>1){
					fl3value=sheet.getCell(2, i-1).getContents().trim();
				}
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
					String bb = null;
					if(sheet.getCell(6, i).getContents() != null){
						String value = null;
						seltext=sheet.getCell(6, i).getContents().trim();
						value = seltext+"; ";
						for(int j = 1; j <= 10; j++){
							//String aa = sheet.getCell(6+j, i).getContents().trim();
							//value = aa+";";
							if(sheet.getCell(6+j, i).getContents() != null && !"".equals(sheet.getCell(6+j, i).getContents())){
								value = value + sheet.getCell(6+j, i).getContents().trim()+"; ";
								
								//System.out.println("qqqq: "+aa);
								bb = value;
								//System.out.println("qqqq: "+value);
							}
							
						}
					}
					seltext = bb; 
								
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
					
					/*GenExcelSqlUtil genChangjiaExcelSql = new GenExcelSqlUtil();
					int rsNume = genChangjiaExcelSql.getMsgrpIdNum(fl1,fl2,fl3);*/
					/*if(rsNume == 1){
						sql2.append("insert into mshpggmb_grp(owner_id,owner_name,flag,org_id,op_id,op_dpt,op_ip) values((select m3.id from mshpfl3 m3 where m3.name = '"+fl3+"' and m3.shpfl2_id = (select m2.id from mshpfl2 m2 where m2.name ='"+fl2+"' and m2.shpfl1_id=(select m1.id from mshpfl1 m1 where m1.name='"+fl1+"'))),'shpfl3_id',1,0,0,3,'0:0:0:0:0:0:0:1');\r\n");
					}else if(rsNume > 1){
						sql2.append("insert into mshpggmb_grp(owner_id,owner_name,flag,org_id,op_id,op_dpt,op_ip) values((select m3.id from mshpfl3 m3 where m3.name = '"+fl3+"' and m3.shpfl2_id = (select m2.id from mshpfl2 m2 where m2.name ='"+fl2+"' and m2.shpfl1_id=(select m1.id from mshpfl1 m1 where m1.name='"+fl1+"'))),'shpfl3_id',0,(select mgp.id from mshpggmb_grp mgp where mgp.owner_id = (select m3.id from mshpfl3 m3 where m3.name= '"+fl3+"' and m3.shpfl2_id = (select m2.id from mshpfl2 m2 where m2.name ='"+fl2+"' and m2.shpfl1_id=(select m1.id from mshpfl1 m1 where m1.name='"+fl1+"'))) and mgp.flag = 1),0,3,'0:0:0:0:0:0:0:1');\r\n");
					}*/
					
					if(!fl3value.equals(fl3)){
						sql2.append("insert into mshpggmb_grp(owner_id,owner_name,flag,org_id,op_id,op_dpt,op_ip) values((select m3.id from mshpfl3 m3 where m3.name = '"+fl3+"' and m3.shpfl2_id = (select m2.id from mshpfl2 m2 where m2.name ='"+fl2+"' and m2.shpfl1_id=(select m1.id from mshpfl1 m1 where m1.name='"+fl1+"'))),'shpfl3_id',1,0,0,3,'0:0:0:0:0:0:0:1');\r\n");
					}else if(fl3value.equals(fl3)){
						sql2.append("insert into mshpggmb_grp(owner_id,owner_name,flag,org_id,op_id,op_dpt,op_ip) values((select m3.id from mshpfl3 m3 where m3.name = '"+fl3+"' and m3.shpfl2_id = (select m2.id from mshpfl2 m2 where m2.name ='"+fl2+"' and m2.shpfl1_id=(select m1.id from mshpfl1 m1 where m1.name='"+fl1+"'))),'shpfl3_id',0,(select mgp.id from mshpggmb_grp mgp where mgp.owner_id = (select m3.id from mshpfl3 m3 where m3.name= '"+fl3+"' and m3.shpfl2_id = (select m2.id from mshpfl2 m2 where m2.name ='"+fl2+"' and m2.shpfl1_id=(select m1.id from mshpfl1 m1 where m1.name='"+fl1+"'))) and mgp.flag = 1),0,3,'0:0:0:0:0:0:0:1');\r\n");
					}
					
					//sql.append("insert into mshpggmb(grp_id,rownum,mblk_id,shpfl3_id,ltext,lkind,bitian,jdisp,selitem,deftext,op_ip) values((select mgp.id from mshpggmb_grp mgp where mgp.owner_id = (select m3.id from mshpfl3 m3 where m3.name = '"+fl3+"' and m3.shpfl2_id = (select m2.id from mshpfl2 m2 where m2.name ='"+fl2+"' and m2.shpfl1_id=(select m1.id from mshpfl1 m1 where m1.name='"+fl1+"'))) and mgp.flag = 1),(select count(mgp.id)+1 from mshpggmb mgp where mgp.shpfl3_id = (select m3.id from mshpfl3 m3 where m3.name = '"+fl3+"' and m3.shpfl2_id = (select m2.id from mshpfl2 m2 where m2.name ='"+fl2+"' and m2.shpfl1_id=(select m1.id from mshpfl1 m1 where m1.name='"+fl1+"'))) and mgp.grp_id = (select mgp.id from mshpggmb_grp mgp where mgp.owner_id = (select m3.id from mshpfl3 m3 where m3.name = '"+fl3+"' and m3.shpfl2_id = (select m2.id from mshpfl2 m2 where m2.name ='"+fl2+"' and m2.shpfl1_id=(select m1.id from mshpfl1 m1 where m1.name='"+fl1+"'))) and mgp.flag = 1) and mgp.flag = 1),0,(select m3.id from mshpfl3 m3 where m3.name = '"+fl3+"' and m3.shpfl2_id = (select m2.id from mshpfl2 m2 where m2.name ='"+fl2+"' and m2.shpfl1_id=(select m1.id from mshpfl1 m1 where m1.name='"+fl1+"'))),'"+ltext+"',"+lkind+","+bitian+","+jdisp+","+selitem+",'"+deftext+"','0:0:0:0:0:0:0:1','"+seltext+"');\r\n");
					sql2.append("insert into mshpggmb(grp_id,rownum,mblk_id,shpfl3_id,ltext,lkind,bitian,jdisp,selitem,deftext,op_ip,seltext) values((select mgp.id from mshpggmb_grp mgp where mgp.owner_id = (select m3.id from mshpfl3 m3 where m3.name = '"+fl3+"' and m3.shpfl2_id = (select m2.id from mshpfl2 m2 where m2.name ='"+fl2+"' and m2.shpfl1_id=(select m1.id from mshpfl1 m1 where m1.name='"+fl1+"'))) and mgp.flag = 1),(select count(mgp.id)+1 from mshpggmb mgp where mgp.shpfl3_id = (select m3.id from mshpfl3 m3 where m3.name = '"+fl3+"' and m3.shpfl2_id = (select m2.id from mshpfl2 m2 where m2.name ='"+fl2+"' and m2.shpfl1_id=(select m1.id from mshpfl1 m1 where m1.name='"+fl1+"'))) and mgp.grp_id = (select mgp.id from mshpggmb_grp mgp where mgp.owner_id = (select m3.id from mshpfl3 m3 where m3.name = '"+fl3+"' and m3.shpfl2_id = (select m2.id from mshpfl2 m2 where m2.name ='"+fl2+"' and m2.shpfl1_id=(select m1.id from mshpfl1 m1 where m1.name='"+fl1+"'))) and mgp.flag = 1) and mgp.flag = 1),0,(select m3.id from mshpfl3 m3 where m3.name = '"+fl3+"' and m3.shpfl2_id = (select m2.id from mshpfl2 m2 where m2.name ='"+fl2+"' and m2.shpfl1_id=(select m1.id from mshpfl1 m1 where m1.name='"+fl1+"'))),'"+ltext+"',"+lkind+","+bitian+","+jdisp+","+selitem+",'"+deftext+"','0:0:0:0:0:0:0:1','"+seltext+"');\r\n");

					//sql.append("insert into mshpggmb(shpfl3_id,ltext,lkind,bitian,jdisp,selitem,deftext,op_ip,seltext) values((select m3.id from mshpfl3 m3 where m3.name = '"+fl3+"' and m3.shpfl2_id = (select m2.id from mshpfl2 m2 where m2.name ='"+fl2+"' and m2.shpfl1_id=(select m1.id from mshpfl1 m1 where m1.name='"+fl1+"'))),'"+ltext+"',"+lkind+","+bitian+","+jdisp+","+selitem+",'"+deftext+"','0:0:0:0:0:0:0:1','"+seltext+"');\r\n");
					System.out.println("��Ʒ����:"+fl1+" ��Ʒ����:"+fl2+" ��ƷС��:"+fl3+" ������������:"+ltext+" ��Ʒ����:"+jdisp+" Ĭ��ֵ:"+deftext+" ɸѡֵ:"+seltext);
					//sql.append("insert into mquxian(code,name,city_id) values('"+code+"','"+area+"'"+buqi+", (select id from mshi where name='"+city+"' and sh_id = (select id from msheng where name='"+province+"') limit 1));\r\n");
				
				}
				
				}
			FileWriter fw = new FileWriter(saveSqlf);//�����sql�ļ�
			fw.write(sql2.toString());
			//fw.write(sql.toString());//д������
			fw.close();//�ر�
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

}
