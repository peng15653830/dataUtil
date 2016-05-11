package com.vjd.dtutil;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;

import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

/**
 * V机电产品二级、三级分类初始化数据程序
 * 要保证word文档的格式规范
 * 增加补齐和空行功能
 *
 * @author 陈清水
 */
public class GenWordSql {

    private static String readFile = "/home/xupeng/workspace/dataUtil/src/fenlei4-29.docx"; //要读取的V机电分类文件（word文档）
    private static String saveSqlf = "/home/xupeng/workspace/dataUtil/src/cpfl3.sql"; //要保存的V机电分类SQL（二级、三级）
    private static int maxLen = 22;//名称最大长度

    public static void main(String[] args) {
        try {
            FileInputStream in = new FileInputStream(new File(readFile));
            XWPFDocument doc = new XWPFDocument(in);
            XWPFWordExtractor ext = new XWPFWordExtractor(doc);
            String cont = ext.getText();//word内容
            String[] yjfl = {"一、", "二、", "三、", "四、", "五、", "六、", "七、"};//一级分类
            String[] yjbt = new String[yjfl.length];//一级标题
            String code = null;//code
            StringBuilder sbej = new StringBuilder();//二级sql
            StringBuilder sbsj = new StringBuilder();//三级sql
            int nxtp = 0;//下一个位置
            for (int i = 0; i < yjfl.length; i++) {
                int a = cont.indexOf(yjfl[i]);
                int b = cont.indexOf("\n", a);//换行符
                yjbt[i] = cont.substring(a + 2, b);
                if (i < yjfl.length - 1) {
                    nxtp = cont.indexOf(yjfl[i + 1]);
                } else
                    nxtp = cont.indexOf("备注：");//最后备注
                a = cont.indexOf("：", b);
                //System.out.println("一级："+yjbt[i]);
                int j = 0;//二级序号
                while (a < nxtp) {
                    j = j + 1;
                    if (j <= 9)
                        code = "0" + j;//编号不足前面补0
                    else
                        code = "" + j;
                    String ejbt = cont.substring(b + 1, a);//二级标题
                    int len = maxLen - 2 * ejbt.length();
                    String buqi = " ";
                    for (int k = 1; k < len; k++)
                        buqi = buqi + " ";
                    sbej.append("insert into mshpfl2 (code, name, shpfl1_id) values('" + code + "', '" + ejbt + "'" + buqi + ", (select id from mshpfl1 where name='" + yjbt[i] + "' limit 1));\n");
                    b = cont.indexOf("\n", a);
                    String sjbt = cont.substring(a + 1, b).replaceFirst("。", "");//三级标题,去掉句号
                    a = cont.indexOf("：", b);
                    if (sjbt.equals(""))
                        continue;//没有三级不往下处理
                    String[] sjbts = sjbt.split("，");
                    for (int k = 0; k < sjbts.length; k++) {//循环三级标题
                        if (k < 9)
                            code = "0" + (k + 1);
                        else
                            code = "" + (k + 1);
                        len = maxLen - 2 * sjbts[k].length();
                        buqi = " ";
                        String buqi2 = " ";
                        for (int m = 0; m < sjbts[k].length(); m++) {
                            if (sjbts[k].charAt(m) < 128) //英文字符
                                len = len + 1;
                        }
                        for (int m = 1; m < len; m++)
                            buqi = buqi + " ";
                        len = 10 - 2 * yjbt[i].length();
                        for (int m = 0; m < yjbt[i].length(); m++) {
                            if (yjbt[i].charAt(m) < 128) //英文字符
                                len = len + 1;
                        }
                        for (int m = 1; m < len; m++)
                            buqi2 = buqi2 + " ";
                        sbsj.append("insert into mshpfl3 (code, name, shpfl1_id, shpfl2_id) values('" + code + "', '" + sjbts[k] + "'" + buqi + ", (select id from mshpfl1 where name='" + yjbt[i] + "' limit 1)" + buqi2 + ", (select id from mshpfl2 where name='" + ejbt + "' limit 1));\n");
                    }
                    sbsj.append("\r\n");//从新编号时补空行
                    //System.out.println("二级："+ejbt);
                    //System.out.println("三级："+sjbt);
                }
                sbej.append("\r\n");//从新编号时补空行
            }
            FileWriter fw = new FileWriter(saveSqlf);//保存的sql文件
            sbej.append(sbsj);
            fw.write(sbej.toString());//写入数据
            fw.close();//关闭
            ext.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}