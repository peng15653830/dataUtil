package com.vjd.dtutil;

import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.InputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;

import jxl.Sheet;
import jxl.Workbook;

public class GenChangjiaExcelSql {
	
	private static Connection conn = null;
	private static Statement stmt = null;
	private static Statement stmt2 = null;
	private static ResultSet rs = null;
	private static ResultSet rs2 = null;
	//private static String url = "jdbc:oracle:thin:@172.16.5.189:1521:cr1221";
	//private static String url = "jdbc:mysql://127.0.0.1:3306/mydb";
	private static String url = "jdbc:postgresql://192.168.1.201:5432/vjd_v1.1";
	//private static String url = PropertiesUtil.getValue("Database.URL");
	private static String username = "vjidian";
	//private static String username = PropertiesUtil.getValue("Database.UserName");
	private static String password = "vjidian";	
	//private static String password = PropertiesUtil.getValue("Database.Password");
		
	public  static Connection getServiceNumManager(){		
		try {
			//DriverManager.registerDriver(new oracle.jdbc.OracleDriver());
			//DriverManager.registerDriver(new com.mysql.jdbc.Driver());
			DriverManager.registerDriver(new org.postgresql.Driver());
			conn = DriverManager.getConnection(url, username, password);
			return conn;
		} catch (SQLException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return null; 		
    }
	
	/**
     * 
	 * ����ʡ���ƣ�ȡ��ʡ����id 
	 * 
	 */  
    public int getProvinceId(String provinceName){         
    	int provinceIdValue = 0;//ʡid	       			 
   			   try {    				  
	   			 //String province=null;//ʡ����
   				 String province=provinceName;//ʡ����
	   			 String city = null;//����
	   			 String area = null;//���� 	   			 
   				   
   				 conn = getServiceNumManager(); 	     
				 conn.setAutoCommit(false); 		    
   			     stmt = conn.createStatement(); 
   			     //rs = stmt.executeQuery("select *  from cre_doc ");
   			     rs = stmt.executeQuery("select id from msheng where name='"+province+"' and flag=1");

   			     while (rs.next()) {
   			       String provinceId = rs.getString("id");    			          				       	       
   			       provinceIdValue = Integer.parseInt(provinceId);     			    
   			     }      			  
   			     conn.commit();    			     
   			     rs.close(); 
   			   }
   			   catch (Exception e) {
   			     System.out.println("error: " + e);
   			     try {
   			       conn.rollback();
   			     }
   			     catch (SQLException sqle) {}
   			   }
   			   finally { 
   				
   			     try {
   			       if (rs != null){
   			         rs.close();
   			       }   			         
   			     }
   			     catch (SQLException sqle) {
   			       System.out.println("SQLState: " + sqle.getSQLState());
   			       System.out.println("SQLErrorCode: �������" + sqle.getErrorCode());
   			       System.out.println("SQLErrorMessage:����������ַ��� " + sqle.toString());
   			     } 

   			     try {
   			       if (stmt != null){
   			         stmt.close();
   			       }
   			     }
   			     catch (SQLException sqle1) {
   			       System.out.println("SQLState: " + sqle1.getSQLState());
   			       System.out.println("SQLErrorCode: �������" + sqle1.getErrorCode());
   			       System.out.println("SQLErrorMessage:����������ַ��� " + sqle1.toString());
   			     } 

   			     try {
   			       if (conn != null)
   			         conn.close();
   			     }
   			     catch (SQLException sqle2) {
   			       System.out.println(sqle2.toString());
   			       System.out.println(sqle2.getSQLState());
   			       System.out.println(sqle2.getErrorCode());
   			     } 
   			   }
			return provinceIdValue;    			
   	}
    
    /**
     * 
	 * ���������ƺ�ʡid��ȡ��������id 
	 * 
	 */  
    public int getCityId(String cityName,int provinceId){         
    	int cityIdValue = 0;//��id	       			 
   			   try {    				  
	   			 //String province=null;//ʡ����
   				 String city=cityName;//ʡ����   			 
   				   
   				 conn = getServiceNumManager(); 	     
				 conn.setAutoCommit(false); 		    
   			     stmt = conn.createStatement(); 
   			     //rs = stmt.executeQuery("select *  from cre_doc ");
   			     rs = stmt.executeQuery("select id from mshi where name='"+city+"' and sh_id='"+provinceId+"' and flag=1");

   			     while (rs.next()) {  			          				       	       
   			       cityIdValue = Integer.parseInt(rs.getString("id"));     			    
   			     }      			  
   			     conn.commit();    			     
   			     rs.close(); 
   			   }
   			   catch (Exception e) {
   			     System.out.println("error: " + e);
   			     try {
   			       conn.rollback();
   			     }
   			     catch (SQLException sqle) {}
   			   }
   			   finally { 
   				
   			     try {
   			       if (rs != null){
   			         rs.close();
   			       }   			         
   			     }
   			     catch (SQLException sqle) {
   			       System.out.println("SQLState: " + sqle.getSQLState());
   			       System.out.println("SQLErrorCode: �������" + sqle.getErrorCode());
   			       System.out.println("SQLErrorMessage:����������ַ��� " + sqle.toString());
   			     } 

   			     try {
   			       if (stmt != null){
   			         stmt.close();
   			       }
   			     }
   			     catch (SQLException sqle1) {
   			       System.out.println("SQLState: " + sqle1.getSQLState());
   			       System.out.println("SQLErrorCode: �������" + sqle1.getErrorCode());
   			       System.out.println("SQLErrorMessage:����������ַ��� " + sqle1.toString());
   			     } 

   			     try {
   			       if (conn != null)
   			         conn.close();
   			     }
   			     catch (SQLException sqle2) {
   			       System.out.println(sqle2.toString());
   			       System.out.println(sqle2.getSQLState());
   			       System.out.println(sqle2.getErrorCode());
   			     } 
   			   }
			return cityIdValue;    			
   	}
    
    /**
     * 
	 * ������������,��id,ȡ����������id 
	 * 
	 */  
    public int getAreaId(String areaName,int cityId){         
    	int areaIdValue = 0;//����id	       			 
   			   try {    				  
   				 String city=areaName;//��������   			 
   				   
   				 conn = getServiceNumManager(); 	     
				 conn.setAutoCommit(false); 		    
   			     stmt = conn.createStatement(); 
   			     rs = stmt.executeQuery("select id from mquxian where name='"+city+"' and city_id='"+cityId+"' and flag=1");

   			     while (rs.next()) {  			          				       	       
   			    	areaIdValue = Integer.parseInt(rs.getString("id"));     			    
   			     }      			  
   			     conn.commit();    			     
   			     rs.close(); 
   			   }
   			   catch (Exception e) {
   			     System.out.println("error: " + e);
   			     try {
   			       conn.rollback();
   			     }
   			     catch (SQLException sqle) {}
   			   }
   			   finally { 
   				
   			     try {
   			       if (rs != null){
   			         rs.close();
   			       }   			         
   			     }
   			     catch (SQLException sqle) {
   			       System.out.println("SQLState: " + sqle.getSQLState());
   			       System.out.println("SQLErrorCode: �������" + sqle.getErrorCode());
   			       System.out.println("SQLErrorMessage:����������ַ��� " + sqle.toString());
   			     } 

   			     try {
   			       if (stmt != null){
   			         stmt.close();
   			       }
   			     }
   			     catch (SQLException sqle1) {
   			       System.out.println("SQLState: " + sqle1.getSQLState());
   			       System.out.println("SQLErrorCode: �������" + sqle1.getErrorCode());
   			       System.out.println("SQLErrorMessage:����������ַ��� " + sqle1.toString());
   			     } 

   			     try {
   			       if (conn != null)
   			         conn.close();
   			     }
   			     catch (SQLException sqle2) {
   			       System.out.println(sqle2.toString());
   			       System.out.println(sqle2.getSQLState());
   			       System.out.println(sqle2.getErrorCode());
   			     } 
   			   }
			return areaIdValue;    			
   	}
    
    /**
     * 
	 * ���ݷֹ�˾����,ȡ�÷ֹ�˾����id 
	 * 
	 */  
    public int getFengongsiId(String fengongsiName){         
    	int fengongsiIdValue = 0;//�ֹ�˾id	       			 
   			   try { 
   				 conn = getServiceNumManager(); 	     
				 conn.setAutoCommit(false); 		    
   			     stmt = conn.createStatement(); 
   			     rs = stmt.executeQuery("select id from mfengs where name='"+fengongsiName+"' and flag=1");

   			     while (rs.next()) {  			          				       	       
   			    	fengongsiIdValue = Integer.parseInt(rs.getString("id"));     			    
   			     }      			  
   			     conn.commit();    			     
   			     rs.close(); 
   			   }
   			   catch (Exception e) {
   			     System.out.println("error: " + e);
   			     try {
   			       conn.rollback();
   			     }
   			     catch (SQLException sqle) {}
   			   }
   			   finally { 
   				
   			     try {
   			       if (rs != null){
   			         rs.close();
   			       }   			         
   			     }
   			     catch (SQLException sqle) {
   			       System.out.println("SQLState: " + sqle.getSQLState());
   			       System.out.println("SQLErrorCode: �������" + sqle.getErrorCode());
   			       System.out.println("SQLErrorMessage:����������ַ��� " + sqle.toString());
   			     } 

   			     try {
   			       if (stmt != null){
   			         stmt.close();
   			       }
   			     }
   			     catch (SQLException sqle1) {
   			       System.out.println("SQLState: " + sqle1.getSQLState());
   			       System.out.println("SQLErrorCode: �������" + sqle1.getErrorCode());
   			       System.out.println("SQLErrorMessage:����������ַ��� " + sqle1.toString());
   			     } 

   			     try {
   			       if (conn != null)
   			         conn.close();
   			     }
   			     catch (SQLException sqle2) {
   			       System.out.println(sqle2.toString());
   			       System.out.println(sqle2.getSQLState());
   			       System.out.println(sqle2.getErrorCode());
   			     } 
   			   }
			return fengongsiIdValue;    			
   	}
    
    /**
     * 
	 * ����Ա������,�ֹ�˾id,ȡ��Ա������id 
	 * 
	 */  
    public int getFengongsiYuangongId(String fengongsiYuangongName,int fengongsiid){         
    	int fengongsiYuangongIdValue = 0;//�ֹ�˾Ա��id	       			 
   			   try { 
   				 conn = getServiceNumManager(); 	     
				 conn.setAutoCommit(false); 		    
   			     stmt = conn.createStatement(); 
   			     rs = stmt.executeQuery("select id from mfgsyg where name='"+fengongsiYuangongName+"' and fgs_id='"+fengongsiid+"' and flag=1");

   			     while (rs.next()) {  			          				       	       
   			    	fengongsiYuangongIdValue = Integer.parseInt(rs.getString("id"));     			    
   			     }      			  
   			     conn.commit();    			     
   			     rs.close(); 
   			   }
   			   catch (Exception e) {
   			     System.out.println("error: " + e);
   			     try {
   			       conn.rollback();
   			     }
   			     catch (SQLException sqle) {}
   			   }
   			   finally { 
   				
   			     try {
   			       if (rs != null){
   			         rs.close();
   			       }   			         
   			     }
   			     catch (SQLException sqle) {
   			       System.out.println("SQLState: " + sqle.getSQLState());
   			       System.out.println("SQLErrorCode: �������" + sqle.getErrorCode());
   			       System.out.println("SQLErrorMessage:����������ַ��� " + sqle.toString());
   			     } 

   			     try {
   			       if (stmt != null){
   			         stmt.close();
   			       }
   			     }
   			     catch (SQLException sqle1) {
   			       System.out.println("SQLState: " + sqle1.getSQLState());
   			       System.out.println("SQLErrorCode: �������" + sqle1.getErrorCode());
   			       System.out.println("SQLErrorMessage:����������ַ��� " + sqle1.toString());
   			     } 

   			     try {
   			       if (conn != null)
   			         conn.close();
   			     }
   			     catch (SQLException sqle2) {
   			       System.out.println(sqle2.toString());
   			       System.out.println(sqle2.getSQLState());
   			       System.out.println(sqle2.getErrorCode());
   			     } 
   			   }
			return fengongsiYuangongIdValue;    			
   	}
    
    /**
     * 
	 * ������Ʒ�������ƣ�ȡ������id 
	 * 
	 */  
    public int getShangpinleixingId(String leixingName){         
    	int shangpinleixingIdValue = 0;//ʡid	       			 
   			   try { 
   				 conn = getServiceNumManager(); 	     
				 conn.setAutoCommit(false); 		    
   			     stmt = conn.createStatement(); 
   			     rs = stmt.executeQuery("select id from mshpfl1 where name='"+leixingName+"' and flag=1");

   			     while (rs.next()) { 			          				       	       
   			    	 shangpinleixingIdValue = Integer.parseInt(rs.getString("id"));     			    
   			     }      			  
   			     conn.commit();    			     
   			     rs.close(); 
   			   }
   			   catch (Exception e) {
   			     System.out.println("error: " + e);
   			     try {
   			       conn.rollback();
   			     }
   			     catch (SQLException sqle) {}
   			   }
   			   finally { 
   				
   			     try {
   			       if (rs != null){
   			         rs.close();
   			       }   			         
   			     }
   			     catch (SQLException sqle) {
   			       System.out.println("SQLState: " + sqle.getSQLState());
   			       System.out.println("SQLErrorCode: �������" + sqle.getErrorCode());
   			       System.out.println("SQLErrorMessage:����������ַ��� " + sqle.toString());
   			     } 

   			     try {
   			       if (stmt != null){
   			         stmt.close();
   			       }
   			     }
   			     catch (SQLException sqle1) {
   			       System.out.println("SQLState: " + sqle1.getSQLState());
   			       System.out.println("SQLErrorCode: �������" + sqle1.getErrorCode());
   			       System.out.println("SQLErrorMessage:����������ַ��� " + sqle1.toString());
   			     } 

   			     try {
   			       if (conn != null)
   			         conn.close();
   			     }
   			     catch (SQLException sqle2) {
   			       System.out.println(sqle2.toString());
   			       System.out.println(sqle2.getSQLState());
   			       System.out.println(sqle2.getErrorCode());
   			     } 
   			   }
			return shangpinleixingIdValue;    			
   	}
    
    /**
     * 
	 * ��ӳ�����Ϣ 
	 * 
	 */  
    public int addChangjia(int provinceId,int cityId,int areaId,int fengongsiid,int fengongsiyuangongid,int shangpinleixingid,String changjianame,String jingyingfanwei,String urlvalue,String faxvalue,String phonevalue,String emailvalue,String addr4value,String lnamevalue,String lphonevalue,String lemailvalue,String bunessname,String bunessphonevalue,String bunessemailvalue){         
    	int insertValue = 0;//������ݽ��	       			 
   			   try { 
   				 conn = getServiceNumManager(); 	     
				 conn.setAutoCommit(false); 		    
   			     stmt = conn.createStatement(); 
   			     String sql = "INSERT INTO mchangj (fgs_id, fgsyg_id, shpfl1_id, name, jyfw, url, fax, phone, email, addr1, addr2, addr3, addr4, lname, mobile, lemail, bname, bphone, bemail, passwd, flag, org_id, op_id, op_dpt, op_ip) VALUES ('"+fengongsiid+"', '"+fengongsiyuangongid+"', '"+shangpinleixingid+"', '"+changjianame+"', '"+jingyingfanwei+"', '"+urlvalue+"', '"+faxvalue+"', '"+phonevalue+"', '"+emailvalue+"', '"+provinceId+"', '"+cityId+"', '"+areaId+"', '"+addr4value+"', '"+lnamevalue+"', '"+lphonevalue+"', '"+lemailvalue+"', '"+bunessname+"', '"+bunessphonevalue+"', '"+bunessemailvalue+"', 'e10adc3949ba59abbe56e057f20f883e', '1', '0', '"+fengongsiid+"', '1', '0:0:0:0:0:0:0:1')";
   			     insertValue = stmt.executeUpdate(sql);    			  
   			     conn.commit();    			     
   			     rs.close(); 
   			   }
   			   catch (Exception e) {
   			     System.out.println("error: " + e);
   			     try {
   			       conn.rollback();
   			     }
   			     catch (SQLException sqle) {}
   			   }
   			   finally { 
   				
   			     try {
   			       if (rs != null){
   			         rs.close();
   			       }   			         
   			     }
   			     catch (SQLException sqle) {
   			       System.out.println("SQLState: " + sqle.getSQLState());
   			       System.out.println("SQLErrorCode: �������" + sqle.getErrorCode());
   			       System.out.println("SQLErrorMessage:����������ַ��� " + sqle.toString());
   			     } 

   			     try {
   			       if (stmt != null){
   			         stmt.close();
   			       }
   			     }
   			     catch (SQLException sqle1) {
   			       System.out.println("SQLState: " + sqle1.getSQLState());
   			       System.out.println("SQLErrorCode: �������" + sqle1.getErrorCode());
   			       System.out.println("SQLErrorMessage:����������ַ��� " + sqle1.toString());
   			     } 

   			     try {
   			       if (conn != null)
   			         conn.close();
   			     }
   			     catch (SQLException sqle2) {
   			       System.out.println(sqle2.toString());
   			       System.out.println(sqle2.getSQLState());
   			       System.out.println(sqle2.getErrorCode());
   			     } 
   			   }
			return insertValue;    			
   	}
    
    /**
     * 
	 * ��ȡ����ģ��excel�ļ���Ϣ���� 
	 * 
	 */
    public int readExcelChangjia(String readFile,String saveSqlf){
    	int readValue = 0;
    	try {
			Workbook rwb = null;
			// ����������
			InputStream stream = new FileInputStream(readFile);
			// ��ȡExcel�ļ�����
			rwb = Workbook.getWorkbook(stream);
			// ��ȡ�ļ���ָ�������� Ĭ�ϵĵ�һ��
			Sheet sheet = rwb.getSheet(0);
			int rowct = sheet.getRows();
			
			String fengognsiname = null;//�ֹ�˾����
			String fengongsiyuangongname = null;//�ֹ�˾Ա������
			String shangpinleixing = null;//��Ӫ��Ʒ����
			String changjianame = null;//��������
			String jingyingfanwei = null;//��Ӫ��Χ
			String urlvalue = null;//��ַ
			String faxnumber = null;//����
			String phonenumber = null;//�绰
			String mailaddre = null;//����
			String fl1=null;//ʡ
			String fl2 = null;//��
			String fl3 = null;//����
			String xxaddre = null;//��ϸ��ַ
			String linkname = null;//��ϵ������
			String linkphone = null;//��ϵ�˵绰
			String linkmail = null;//��ϵ������
			String bunesslinkname = null;//������ϵ������
			String bunesslinkphone = null;//������ϵ�˵绰
			String bunesslinkmail = null;//������ϵ������
			StringBuilder sql = new StringBuilder();
			for (int i = 1; i < rowct; i++) {
				
				fengognsiname=sheet.getCell(0, i).getContents().trim();//�ֹ�˾����
				fengongsiyuangongname=sheet.getCell(1, i).getContents().trim();//�ֹ�˾Ա������
				shangpinleixing=sheet.getCell(2, i).getContents().trim();//��Ӫ��Ʒ����
				changjianame=sheet.getCell(3, i).getContents().trim();//��������
				jingyingfanwei=sheet.getCell(4, i).getContents().trim();//��Ӫ��Χ
				urlvalue=sheet.getCell(5, i).getContents().trim();//��ַ
				faxnumber=sheet.getCell(6, i).getContents().trim();//����
				phonenumber=sheet.getCell(7, i).getContents().trim();//�绰
				mailaddre=sheet.getCell(8, i).getContents().trim();//����
				
				fl1 =sheet.getCell(9, i).getContents().trim();//ʡ
				fl2 =sheet.getCell(10, i).getContents().trim();//��
				fl3 =sheet.getCell(11, i).getContents().trim();//����				
				
				xxaddre=sheet.getCell(12, i).getContents().trim();//��ϸ��ַ
				linkname=sheet.getCell(13, i).getContents().trim();//��ϵ������
				linkphone=sheet.getCell(14, i).getContents().trim();//��ϵ�˵绰
				linkmail=sheet.getCell(15, i).getContents().trim();//��ϵ������
				bunesslinkname=sheet.getCell(16, i).getContents().trim();//������ϵ��
				bunesslinkphone=sheet.getCell(17, i).getContents().trim();//������ϵ�˵绰
				bunesslinkmail=sheet.getCell(18, i).getContents().trim();//������ϵ������
				
				GenChangjiaExcelSql genChangjiaExcelSql2 = new GenChangjiaExcelSql();
				int fl1aa = genChangjiaExcelSql2.getProvinceId(fl1);
				int fl2bb = genChangjiaExcelSql2.getCityId(fl2,fl1aa);
				int fl3cc = genChangjiaExcelSql2.getAreaId(fl3, fl2bb);
				int fengongsiidvalue = genChangjiaExcelSql2.getFengongsiId(fengognsiname);
				int fengongsiyuangongidvalue = genChangjiaExcelSql2.getFengongsiYuangongId(fengongsiyuangongname, fengongsiidvalue);
				int shangpinleixingidvalue = genChangjiaExcelSql2.getShangpinleixingId(shangpinleixing);
				
				//��ӳ���������Ϣ
				if(fengognsiname != null && !"".equals(fengognsiname)){//�ֹ�˾��Ϊ�գ��������
					
					//��ѯ���������
					int rsCountNum=genChangjiaExcelSql2.getChangjia(fl1aa,fl2bb,fl3cc,fengongsiidvalue,fengongsiyuangongidvalue,shangpinleixingidvalue,changjianame);
					if(rsCountNum == 1){
						//�޸ĳ���������Ϣ
						genChangjiaExcelSql2.updateChangjia(fl1aa, fl2bb, fl3cc, fengongsiidvalue, fengongsiyuangongidvalue, shangpinleixingidvalue, changjianame, jingyingfanwei, urlvalue, faxnumber, phonenumber, mailaddre, xxaddre, linkname, linkphone, linkmail, bunesslinkname, bunesslinkphone, bunesslinkmail);
					}else if(rsCountNum <1){
						//��ӳ���������Ϣ
						genChangjiaExcelSql2.addChangjia(fl1aa,fl2bb,fl3cc,fengongsiidvalue,fengongsiyuangongidvalue,shangpinleixingidvalue,changjianame,jingyingfanwei,urlvalue,faxnumber,phonenumber,mailaddre,xxaddre,linkname,linkphone,linkmail,bunesslinkname,bunesslinkphone,bunesslinkmail);
					}					
					sql.append("�ļ������������");
					System.out.println("ʡ:"+fl1+" ��:"+fl2+" ����:"+fl3+" �����ֹ�˾����:"+fengognsiname+" �ֹ�˾Ա������:"+fengongsiyuangongname+" ��Ӫ��Ʒ����:"+shangpinleixing+" ��������:"+changjianame+" ��Ӫ��Χ:"+jingyingfanwei+" ��ַ:"+urlvalue+" ����:"+faxnumber+" �绰����:"+phonenumber+" ����:"+mailaddre+" ��ϸ��ַ:"+xxaddre+" ��ϵ������:"+linkname+" ��ϵ�˵绰:"+linkphone+" ��ϵ������:"+linkmail+" ������ϵ��:"+bunesslinkname+" ������ϵ�˵绰:"+bunesslinkphone+" ������ϵ������:"+bunesslinkmail);
					//sql.append("insert into mquxian(code,name,city_id) values('"+code+"','"+area+"'"+buqi+", (select id from mshi where name='"+city+"' and sh_id = (select id from msheng where name='"+province+"') limit 1));\r\n");
				}
			}
			FileWriter fw = new FileWriter(saveSqlf);//�����sql�ļ�
			fw.write(sql.toString());//д������
			fw.close();//�ر�
		} catch (Exception e) {
			e.printStackTrace();
		}
    	return readValue;
    }
    
    /**
     * 
	 * ���ݳ������ƣ��ֹ�˾Ա���������ֹ�˾����Ӫ���ͣ�ʡ���У����ص�ַ  ȡ�ó������� 
	 * 
	 */  
    public int getChangjia(int provinceId,int cityId,int areaId,int fengongsiid,int fengongsiyuangongid,int shangpinleixingid,String changjianame){         
    	int rsCountValue = 0;//��ѯ�������	       			 
   			   try { 
   				 conn = getServiceNumManager(); 	     
				 conn.setAutoCommit(false); 		    
   			     stmt = conn.createStatement(); 
   			     String sqlValue = "select * from mchangj where fgs_id='"+fengongsiid+"' and fgsyg_id='"+fengongsiyuangongid+"' and shpfl1_id='"+shangpinleixingid+"' and addr1='"+provinceId+"' and addr2='"+cityId+"' and addr3='"+areaId+"' and name='"+changjianame+"' and flag=1";
   			     rs = stmt.executeQuery(sqlValue);
   			     
	   			 while (rs.next()) {
	   				rsCountValue = rsCountValue +1; 
	   			 }
   			     
   			     conn.commit();    			     
   			     rs.close(); 
   			   }
   			   catch (Exception e) {
   			     System.out.println("error: " + e);
   			     try {
   			       conn.rollback();
   			     }
   			     catch (SQLException sqle) {}
   			   }
   			   finally { 
   				
   			     try {
   			       if (rs != null){
   			         rs.close();
   			       }   			         
   			     }
   			     catch (SQLException sqle) {
   			       System.out.println("SQLState: " + sqle.getSQLState());
   			       System.out.println("SQLErrorCode: �������" + sqle.getErrorCode());
   			       System.out.println("SQLErrorMessage:����������ַ��� " + sqle.toString());
   			     } 

   			     try {
   			       if (stmt != null){
   			         stmt.close();
   			       }
   			     }
   			     catch (SQLException sqle1) {
   			       System.out.println("SQLState: " + sqle1.getSQLState());
   			       System.out.println("SQLErrorCode: �������" + sqle1.getErrorCode());
   			       System.out.println("SQLErrorMessage:����������ַ��� " + sqle1.toString());
   			     } 

   			     try {
   			       if (conn != null)
   			         conn.close();
   			     }
   			     catch (SQLException sqle2) {
   			       System.out.println(sqle2.toString());
   			       System.out.println(sqle2.getSQLState());
   			       System.out.println(sqle2.getErrorCode());
   			     } 
   			   }
			return rsCountValue;    			
   	}
    
    /**
     * 
	 * �޸ĳ�����Ϣ 
	 * 
	 */  
    public int updateChangjia(int provinceId,int cityId,int areaId,int fengongsiid,int fengongsiyuangongid,int shangpinleixingid,String changjianame,String jingyingfanwei,String urlvalue,String faxvalue,String phonevalue,String emailvalue,String addr4value,String lnamevalue,String lphonevalue,String lemailvalue,String bunessname,String bunessphonevalue,String bunessemailvalue){         
    	int insertValue = 0;//������ݽ��	       			 
   			   try { 
   				 conn = getServiceNumManager(); 	     
				 conn.setAutoCommit(false); 		    
   			     stmt = conn.createStatement(); 
   			     //String sql = "UPDATE mchangj SET fgs_id='"+fengongsiid+"', fgsyg_id='"+fengongsiyuangongid+"', shpfl1_id='"+shangpinleixingid+"', name='"+changjianame+"', jyfw='"+jingyingfanwei+"', url='"+urlvalue+"', fax='"+faxvalue+"', phone='"+phonevalue+"', email='"+emailvalue+"', addr1='"+provinceId+"', addr2='"+cityId+"', addr3='"+areaId+"', addr4='"+addr4value+"', lname='"+lnamevalue+"', mobile='"+lphonevalue+"', lemail='"+lemailvalue+"', bname='"+bunessname+"', bphone='"+bunessphonevalue+"', bemail='"+bunessemailvalue+"' WHERE fgs_id='"+fengongsiid+"' and fgsyg_id='"+fengongsiyuangongid+"' and shpfl1_id='"+shangpinleixingid+"' and addr1='"+provinceId+"' and addr2='"+cityId+"' and addr3='"+areaId+"' and name='"+changjianame+"' and flag=1";
   			     String sql = "UPDATE mchangj SET fgs_id='"+fengongsiid+"', fgsyg_id='"+fengongsiyuangongid+"', shpfl1_id='"+shangpinleixingid+"', name='"+changjianame+"', jyfw='"+jingyingfanwei+"', url='"+urlvalue+"', fax='"+faxvalue+"', phone='"+phonevalue+"', email='"+emailvalue+"', addr1='"+provinceId+"', addr2='"+cityId+"', addr3='"+areaId+"', addr4='"+addr4value+"', lname='"+lnamevalue+"', mobile='"+lphonevalue+"', lemail='"+lemailvalue+"', bname='"+bunessname+"', bphone='"+bunessphonevalue+"', bemail='"+bunessemailvalue+"' WHERE fgs_id='"+fengongsiid+"' and fgsyg_id='"+fengongsiyuangongid+"' and shpfl1_id='"+shangpinleixingid+"' and addr1='"+provinceId+"' and addr2='"+cityId+"' and addr3='"+areaId+"' and name='"+changjianame+"' and flag=1";
   			     insertValue = stmt.executeUpdate(sql);    			  
   			     conn.commit();    			     
   			     rs.close(); 
   			     insertValue = insertValue +1;
   			   }
   			   catch (Exception e) {
   			     System.out.println("error: " + e);
   			     try {
   			       conn.rollback();
   			     }
   			     catch (SQLException sqle) {}
   			   }
   			   finally {    				
   			     try {
   			       if (rs != null){
   			         rs.close();
   			       }   			         
   			     }
   			     catch (SQLException sqle) {
   			       System.out.println("SQLState: " + sqle.getSQLState());
   			       System.out.println("SQLErrorCode: �������" + sqle.getErrorCode());
   			       System.out.println("SQLErrorMessage:����������ַ��� " + sqle.toString());
   			     } 
   			     try {
   			       if (stmt != null){
   			         stmt.close();
   			       }
   			     }
   			     catch (SQLException sqle1) {
   			       System.out.println("SQLState: " + sqle1.getSQLState());
   			       System.out.println("SQLErrorCode: �������" + sqle1.getErrorCode());
   			       System.out.println("SQLErrorMessage:����������ַ��� " + sqle1.toString());
   			     } 
   			     try {
   			       if (conn != null)
   			         conn.close();
   			     }
   			     catch (SQLException sqle2) {
   			       System.out.println(sqle2.toString());
   			       System.out.println(sqle2.getSQLState());
   			       System.out.println(sqle2.getErrorCode());
   			     } 
   			   }
			return insertValue;    			
   	}
	
	public static void main(String[] args) {
		GenChangjiaExcelSql genChangjiaExcelSql = new GenChangjiaExcelSql();
		int aa = genChangjiaExcelSql.getProvinceId("�ӱ�ʡ");
		int bb = genChangjiaExcelSql.getCityId("�ػʵ���",aa);
		int cc = genChangjiaExcelSql.getAreaId("������", bb);
		int fengongsiid = genChangjiaExcelSql.getFengongsiId("�����зֹ�˾����������������������������");
		int fengongsiyuangongid = genChangjiaExcelSql.getFengongsiYuangongId("Ա��[101-1001]=====+", fengongsiid);
		int shangpinleixingid = genChangjiaExcelSql.getShangpinleixingId("��𹤾�");
		
		System.out.println("�ӱ�ʡ����id��"+aa);
		System.out.println("�ػʵ�������id��"+bb);
		System.out.println("����������id��"+cc);
		System.out.println("�����ֹ�˾����id��"+fengongsiid);
		System.out.println("�����ֹ�˾Ա��[101-1001]����id��"+fengongsiyuangongid);
		System.out.println("��𹤾�����id��"+shangpinleixingid);
		
		/*int insertvalue = genChangjiaExcelSql.addChangjia(aa,bb,cc,fengongsiid,fengongsiyuangongid,shangpinleixingid);
		System.out.println("��ӳ��ҽ����"+insertvalue);*/
		
		//��ȡexcel�ļ�����
		genChangjiaExcelSql.readExcelChangjia("d:\\��������¼��ģ��.xls","d:\\changjia.text");
		
		//��ѯ���������
		int rsCountNum=genChangjiaExcelSql.getChangjia(5, 9, 70, 101, 1001, 1, "���Գ�������");
		System.out.println("��ѯ�����������"+rsCountNum);
	}

}
