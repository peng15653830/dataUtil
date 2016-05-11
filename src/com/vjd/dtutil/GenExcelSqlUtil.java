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

public class GenExcelSqlUtil {
	
	private static Connection conn = null;
	private static Statement stmt = null;
	private static Statement stmt2 = null;
	private static ResultSet rs = null;
	private static ResultSet rs2 = null;
	//private static String url = "jdbc:oracle:thin:@172.16.5.189:1521:cr1221";
	//private static String url = "jdbc:mysql://127.0.0.1:3306/mydb";
	private static String url = "jdbc:postgresql://123.103.13.44:5432/vjd_1.0";
	//private static String url = PropertiesUtil.getValue("Database.URL");
	private static String username = "vjidian";
	//private static String username = PropertiesUtil.getValue("Database.UserName");
	private static String password = "1qaz2wsx3edc";	
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
	 * 根据大分类名称，中分类名称，小分类名称确定 mshpggmb_grp 表中对应id数量
	 * 
	 */  
    public int getMsgrpIdNum(String mshpfl1,String mshpfl2,String mshpfl3){ 
    	int rsCountValue = 0;//查询结果行数	
   			   try {  
   				 conn = getServiceNumManager(); 	     
				 conn.setAutoCommit(false); 		    
   			     stmt = conn.createStatement(); 
   			     //rs = stmt.executeQuery("select *  from cre_doc ");
   			     rs = stmt.executeQuery("select count(mgp.id)+1 from mshpggmb_grp mgp where mgp.owner_id = (select m3.id from mshpfl3 m3 where m3.name = '"+mshpfl3+"' and m3.shpfl2_id = (select m2.id from mshpfl2 m2 where m2.name ='"+mshpfl2+"' and m2.shpfl1_id=(select m1.id from mshpfl1 m1 where m1.name='"+mshpfl1+"'))) and mgp.flag = 1");
   			     while (rs.next()) {   			          				       	       
   			    	 //rsCountValue = rsCountValue +1;  
   			    	rsCountValue++;
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
	
	public static void main(String[] args) {
		GenExcelSqlUtil genChangjiaExcelSql = new GenExcelSqlUtil();		
		
		System.out.println("id数量："+genChangjiaExcelSql.getMsgrpIdNum("机电设备","电焊机","对焊机"));
		
	}

}
