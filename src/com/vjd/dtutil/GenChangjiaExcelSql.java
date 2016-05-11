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
	 * 根据省名称，取得省主键id 
	 * 
	 */  
    public int getProvinceId(String provinceName){         
    	int provinceIdValue = 0;//省id	       			 
   			   try {    				  
	   			 //String province=null;//省名称
   				 String province=provinceName;//省名称
	   			 String city = null;//城市
	   			 String area = null;//地区 	   			 
   				   
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
   			       System.out.println("SQLErrorCode: 错误代码" + sqle.getErrorCode());
   			       System.out.println("SQLErrorMessage:错误情况的字符串 " + sqle.toString());
   			     } 

   			     try {
   			       if (stmt != null){
   			         stmt.close();
   			       }
   			     }
   			     catch (SQLException sqle1) {
   			       System.out.println("SQLState: " + sqle1.getSQLState());
   			       System.out.println("SQLErrorCode: 错误代码" + sqle1.getErrorCode());
   			       System.out.println("SQLErrorMessage:错误情况的字符串 " + sqle1.toString());
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
	 * 根据市名称和省id，取得市主键id 
	 * 
	 */  
    public int getCityId(String cityName,int provinceId){         
    	int cityIdValue = 0;//市id	       			 
   			   try {    				  
	   			 //String province=null;//省名称
   				 String city=cityName;//省名称   			 
   				   
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
   			       System.out.println("SQLErrorCode: 错误代码" + sqle.getErrorCode());
   			       System.out.println("SQLErrorMessage:错误情况的字符串 " + sqle.toString());
   			     } 

   			     try {
   			       if (stmt != null){
   			         stmt.close();
   			       }
   			     }
   			     catch (SQLException sqle1) {
   			       System.out.println("SQLState: " + sqle1.getSQLState());
   			       System.out.println("SQLErrorCode: 错误代码" + sqle1.getErrorCode());
   			       System.out.println("SQLErrorMessage:错误情况的字符串 " + sqle1.toString());
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
	 * 根据区县名称,市id,取得区县主键id 
	 * 
	 */  
    public int getAreaId(String areaName,int cityId){         
    	int areaIdValue = 0;//区县id	       			 
   			   try {    				  
   				 String city=areaName;//区县名称   			 
   				   
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
   			       System.out.println("SQLErrorCode: 错误代码" + sqle.getErrorCode());
   			       System.out.println("SQLErrorMessage:错误情况的字符串 " + sqle.toString());
   			     } 

   			     try {
   			       if (stmt != null){
   			         stmt.close();
   			       }
   			     }
   			     catch (SQLException sqle1) {
   			       System.out.println("SQLState: " + sqle1.getSQLState());
   			       System.out.println("SQLErrorCode: 错误代码" + sqle1.getErrorCode());
   			       System.out.println("SQLErrorMessage:错误情况的字符串 " + sqle1.toString());
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
	 * 根据分公司名称,取得分公司主键id 
	 * 
	 */  
    public int getFengongsiId(String fengongsiName){         
    	int fengongsiIdValue = 0;//分公司id	       			 
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
   			       System.out.println("SQLErrorCode: 错误代码" + sqle.getErrorCode());
   			       System.out.println("SQLErrorMessage:错误情况的字符串 " + sqle.toString());
   			     } 

   			     try {
   			       if (stmt != null){
   			         stmt.close();
   			       }
   			     }
   			     catch (SQLException sqle1) {
   			       System.out.println("SQLState: " + sqle1.getSQLState());
   			       System.out.println("SQLErrorCode: 错误代码" + sqle1.getErrorCode());
   			       System.out.println("SQLErrorMessage:错误情况的字符串 " + sqle1.toString());
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
	 * 根据员工名称,分公司id,取得员工主键id 
	 * 
	 */  
    public int getFengongsiYuangongId(String fengongsiYuangongName,int fengongsiid){         
    	int fengongsiYuangongIdValue = 0;//分公司员工id	       			 
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
   			       System.out.println("SQLErrorCode: 错误代码" + sqle.getErrorCode());
   			       System.out.println("SQLErrorMessage:错误情况的字符串 " + sqle.toString());
   			     } 

   			     try {
   			       if (stmt != null){
   			         stmt.close();
   			       }
   			     }
   			     catch (SQLException sqle1) {
   			       System.out.println("SQLState: " + sqle1.getSQLState());
   			       System.out.println("SQLErrorCode: 错误代码" + sqle1.getErrorCode());
   			       System.out.println("SQLErrorMessage:错误情况的字符串 " + sqle1.toString());
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
	 * 根据商品类型名称，取得主键id 
	 * 
	 */  
    public int getShangpinleixingId(String leixingName){         
    	int shangpinleixingIdValue = 0;//省id	       			 
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
   			       System.out.println("SQLErrorCode: 错误代码" + sqle.getErrorCode());
   			       System.out.println("SQLErrorMessage:错误情况的字符串 " + sqle.toString());
   			     } 

   			     try {
   			       if (stmt != null){
   			         stmt.close();
   			       }
   			     }
   			     catch (SQLException sqle1) {
   			       System.out.println("SQLState: " + sqle1.getSQLState());
   			       System.out.println("SQLErrorCode: 错误代码" + sqle1.getErrorCode());
   			       System.out.println("SQLErrorMessage:错误情况的字符串 " + sqle1.toString());
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
	 * 添加厂家信息 
	 * 
	 */  
    public int addChangjia(int provinceId,int cityId,int areaId,int fengongsiid,int fengongsiyuangongid,int shangpinleixingid,String changjianame,String jingyingfanwei,String urlvalue,String faxvalue,String phonevalue,String emailvalue,String addr4value,String lnamevalue,String lphonevalue,String lemailvalue,String bunessname,String bunessphonevalue,String bunessemailvalue){         
    	int insertValue = 0;//添加数据结果	       			 
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
   			       System.out.println("SQLErrorCode: 错误代码" + sqle.getErrorCode());
   			       System.out.println("SQLErrorMessage:错误情况的字符串 " + sqle.toString());
   			     } 

   			     try {
   			       if (stmt != null){
   			         stmt.close();
   			       }
   			     }
   			     catch (SQLException sqle1) {
   			       System.out.println("SQLState: " + sqle1.getSQLState());
   			       System.out.println("SQLErrorCode: 错误代码" + sqle1.getErrorCode());
   			       System.out.println("SQLErrorMessage:错误情况的字符串 " + sqle1.toString());
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
	 * 读取厂家模板excel文件信息数据 
	 * 
	 */
    public int readExcelChangjia(String readFile,String saveSqlf){
    	int readValue = 0;
    	try {
			Workbook rwb = null;
			// 创建输入流
			InputStream stream = new FileInputStream(readFile);
			// 获取Excel文件对象
			rwb = Workbook.getWorkbook(stream);
			// 获取文件的指定工作表 默认的第一个
			Sheet sheet = rwb.getSheet(0);
			int rowct = sheet.getRows();
			
			String fengognsiname = null;//分公司名称
			String fengongsiyuangongname = null;//分公司员工名称
			String shangpinleixing = null;//主营商品类型
			String changjianame = null;//厂家名称
			String jingyingfanwei = null;//经营范围
			String urlvalue = null;//网址
			String faxnumber = null;//传真
			String phonenumber = null;//电话
			String mailaddre = null;//邮箱
			String fl1=null;//省
			String fl2 = null;//市
			String fl3 = null;//区县
			String xxaddre = null;//详细地址
			String linkname = null;//联系人名称
			String linkphone = null;//联系人电话
			String linkmail = null;//联系人邮箱
			String bunesslinkname = null;//商务联系人名称
			String bunesslinkphone = null;//商务联系人电话
			String bunesslinkmail = null;//商务联系人邮箱
			StringBuilder sql = new StringBuilder();
			for (int i = 1; i < rowct; i++) {
				
				fengognsiname=sheet.getCell(0, i).getContents().trim();//分公司名称
				fengongsiyuangongname=sheet.getCell(1, i).getContents().trim();//分公司员工名称
				shangpinleixing=sheet.getCell(2, i).getContents().trim();//主营商品类型
				changjianame=sheet.getCell(3, i).getContents().trim();//厂家名称
				jingyingfanwei=sheet.getCell(4, i).getContents().trim();//经营范围
				urlvalue=sheet.getCell(5, i).getContents().trim();//网址
				faxnumber=sheet.getCell(6, i).getContents().trim();//传真
				phonenumber=sheet.getCell(7, i).getContents().trim();//电话
				mailaddre=sheet.getCell(8, i).getContents().trim();//邮箱
				
				fl1 =sheet.getCell(9, i).getContents().trim();//省
				fl2 =sheet.getCell(10, i).getContents().trim();//市
				fl3 =sheet.getCell(11, i).getContents().trim();//区县				
				
				xxaddre=sheet.getCell(12, i).getContents().trim();//详细地址
				linkname=sheet.getCell(13, i).getContents().trim();//联系人名称
				linkphone=sheet.getCell(14, i).getContents().trim();//联系人电话
				linkmail=sheet.getCell(15, i).getContents().trim();//联系人邮箱
				bunesslinkname=sheet.getCell(16, i).getContents().trim();//商务联系人
				bunesslinkphone=sheet.getCell(17, i).getContents().trim();//商务联系人电话
				bunesslinkmail=sheet.getCell(18, i).getContents().trim();//商务联系人邮箱
				
				GenChangjiaExcelSql genChangjiaExcelSql2 = new GenChangjiaExcelSql();
				int fl1aa = genChangjiaExcelSql2.getProvinceId(fl1);
				int fl2bb = genChangjiaExcelSql2.getCityId(fl2,fl1aa);
				int fl3cc = genChangjiaExcelSql2.getAreaId(fl3, fl2bb);
				int fengongsiidvalue = genChangjiaExcelSql2.getFengongsiId(fengognsiname);
				int fengongsiyuangongidvalue = genChangjiaExcelSql2.getFengongsiYuangongId(fengongsiyuangongname, fengongsiidvalue);
				int shangpinleixingidvalue = genChangjiaExcelSql2.getShangpinleixingId(shangpinleixing);
				
				//添加厂家数据信息
				if(fengognsiname != null && !"".equals(fengognsiname)){//分公司不为空，添加数据
					
					//查询结果数据量
					int rsCountNum=genChangjiaExcelSql2.getChangjia(fl1aa,fl2bb,fl3cc,fengongsiidvalue,fengongsiyuangongidvalue,shangpinleixingidvalue,changjianame);
					if(rsCountNum == 1){
						//修改厂家数据信息
						genChangjiaExcelSql2.updateChangjia(fl1aa, fl2bb, fl3cc, fengongsiidvalue, fengongsiyuangongidvalue, shangpinleixingidvalue, changjianame, jingyingfanwei, urlvalue, faxnumber, phonenumber, mailaddre, xxaddre, linkname, linkphone, linkmail, bunesslinkname, bunesslinkphone, bunesslinkmail);
					}else if(rsCountNum <1){
						//添加厂家数据信息
						genChangjiaExcelSql2.addChangjia(fl1aa,fl2bb,fl3cc,fengongsiidvalue,fengongsiyuangongidvalue,shangpinleixingidvalue,changjianame,jingyingfanwei,urlvalue,faxnumber,phonenumber,mailaddre,xxaddre,linkname,linkphone,linkmail,bunesslinkname,bunesslinkphone,bunesslinkmail);
					}					
					sql.append("文件内容输入测试");
					System.out.println("省:"+fl1+" 市:"+fl2+" 区县:"+fl3+" 所属分公司名称:"+fengognsiname+" 分公司员工名称:"+fengongsiyuangongname+" 主营商品类型:"+shangpinleixing+" 厂家名称:"+changjianame+" 经营范围:"+jingyingfanwei+" 网址:"+urlvalue+" 传真:"+faxnumber+" 电话号码:"+phonenumber+" 邮箱:"+mailaddre+" 详细地址:"+xxaddre+" 联系人名称:"+linkname+" 联系人电话:"+linkphone+" 联系人邮箱:"+linkmail+" 商务联系人:"+bunesslinkname+" 商务联系人电话:"+bunesslinkphone+" 商务联系人邮箱:"+bunesslinkmail);
					//sql.append("insert into mquxian(code,name,city_id) values('"+code+"','"+area+"'"+buqi+", (select id from mshi where name='"+city+"' and sh_id = (select id from msheng where name='"+province+"') limit 1));\r\n");
				}
			}
			FileWriter fw = new FileWriter(saveSqlf);//保存的sql文件
			fw.write(sql.toString());//写入数据
			fw.close();//关闭
		} catch (Exception e) {
			e.printStackTrace();
		}
    	return readValue;
    }
    
    /**
     * 
	 * 根据厂家名称，分公司员工，所属分公司，主营类型，省，市，区县地址  取得厂家数量 
	 * 
	 */  
    public int getChangjia(int provinceId,int cityId,int areaId,int fengongsiid,int fengongsiyuangongid,int shangpinleixingid,String changjianame){         
    	int rsCountValue = 0;//查询结果行数	       			 
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
   			       System.out.println("SQLErrorCode: 错误代码" + sqle.getErrorCode());
   			       System.out.println("SQLErrorMessage:错误情况的字符串 " + sqle.toString());
   			     } 

   			     try {
   			       if (stmt != null){
   			         stmt.close();
   			       }
   			     }
   			     catch (SQLException sqle1) {
   			       System.out.println("SQLState: " + sqle1.getSQLState());
   			       System.out.println("SQLErrorCode: 错误代码" + sqle1.getErrorCode());
   			       System.out.println("SQLErrorMessage:错误情况的字符串 " + sqle1.toString());
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
	 * 修改厂家信息 
	 * 
	 */  
    public int updateChangjia(int provinceId,int cityId,int areaId,int fengongsiid,int fengongsiyuangongid,int shangpinleixingid,String changjianame,String jingyingfanwei,String urlvalue,String faxvalue,String phonevalue,String emailvalue,String addr4value,String lnamevalue,String lphonevalue,String lemailvalue,String bunessname,String bunessphonevalue,String bunessemailvalue){         
    	int insertValue = 0;//添加数据结果	       			 
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
   			       System.out.println("SQLErrorCode: 错误代码" + sqle.getErrorCode());
   			       System.out.println("SQLErrorMessage:错误情况的字符串 " + sqle.toString());
   			     } 
   			     try {
   			       if (stmt != null){
   			         stmt.close();
   			       }
   			     }
   			     catch (SQLException sqle1) {
   			       System.out.println("SQLState: " + sqle1.getSQLState());
   			       System.out.println("SQLErrorCode: 错误代码" + sqle1.getErrorCode());
   			       System.out.println("SQLErrorMessage:错误情况的字符串 " + sqle1.toString());
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
		int aa = genChangjiaExcelSql.getProvinceId("河北省");
		int bb = genChangjiaExcelSql.getCityId("秦皇岛市",aa);
		int cc = genChangjiaExcelSql.getAreaId("海港区", bb);
		int fengongsiid = genChangjiaExcelSql.getFengongsiId("北京市分公司＝＝＝＝＝＝＝＝＝＝＝＝＝＝");
		int fengongsiyuangongid = genChangjiaExcelSql.getFengongsiYuangongId("员工[101-1001]=====+", fengongsiid);
		int shangpinleixingid = genChangjiaExcelSql.getShangpinleixingId("五金工具");
		
		System.out.println("河北省主键id："+aa);
		System.out.println("秦皇岛市主键id："+bb);
		System.out.println("海港区主键id："+cc);
		System.out.println("北京分公司主键id："+fengongsiid);
		System.out.println("北京分公司员工[101-1001]主键id："+fengongsiyuangongid);
		System.out.println("五金工具主键id："+shangpinleixingid);
		
		/*int insertvalue = genChangjiaExcelSql.addChangjia(aa,bb,cc,fengongsiid,fengongsiyuangongid,shangpinleixingid);
		System.out.println("添加厂家结果："+insertvalue);*/
		
		//读取excel文件数据
		genChangjiaExcelSql.readExcelChangjia("d:\\厂家数据录入模板.xls","d:\\changjia.text");
		
		//查询结果数据量
		int rsCountNum=genChangjiaExcelSql.getChangjia(5, 9, 70, 101, 1001, 1, "测试厂家数据");
		System.out.println("查询结果数据量："+rsCountNum);
	}

}
