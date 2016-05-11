package com.vjd.dtutil;

import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.InputStream;

import jxl.Sheet;
import jxl.Workbook;
/**
 * �й����������ݳ�ʼ������
 * Ҫ��֤excel�ĵ��ĸ�ʽ�淶
 * ���Ӳ���Ϳ��й���
 * @author ����ˮ
 *
 */
public class GenExcelSql {
	
	private static String readFile="d:\\�й������������ƻ���.xls"; //Ҫ��ȡ�������ļ���Excel�ĵ���
	private static String saveSqlf="d:\\mquxian.sql"; //Ҫ�������������SQL���ı��ļ���
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
			String province=null;//ʡ����
			String city = null;//����
			String area = null;//����
			String code = null;//����
			StringBuilder sql = new StringBuilder();
			int j=1;//��ʼ��code���
			for (int i = 1; i < rowct; i++) {
				province=sheet.getCell(3, i).getContents().trim();
				city=sheet.getCell(2, i).getContents().trim();
				area=sheet.getCell(1, i).getContents().trim();
				if(i>1){
					if(city.equals(sheet.getCell(2, i-1).getContents()))//��ǰ��������һ�г������
						j=j+1;//���+1
					else //����ȴ��¿�ʼ���
						j=1;
				}
				if(j<10) //���벹0
					code="0"+j;
				else
					code=String.valueOf(j);
				int len = maxLen - 2*area.length();
				String buqi = "";
				for(int k=1;k<len;k++)
					buqi=buqi+" ";
				if(j==1&&i>1)
					sql.append("\r\n");//���±��ʱ������
				if(!city.equals("ʡֱϽ"))//ʡֱϽ���ظ�����
					sql.append("insert into mquxian(code,name,city_id) values('"+code+"','"+area+"'"+buqi+", (select id from mshi where name='"+city+"' limit 1));\r\n");
				else
					sql.append("insert into mquxian(code,name,city_id) values('"+code+"','"+area+"'"+buqi+", (select id from mshi where name='"+city+"' and sh_id = (select id from msheng where name='"+province+"') limit 1));\r\n");
			}
			FileWriter fw = new FileWriter(saveSqlf);//�����sql�ļ�
			fw.write(sql.toString());//д������
			fw.close();//�ر�
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

}
