package org.jpoi.qianjinExcel;

import org.apache.poi.hwpf.usermodel.TableIterator;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * Created by wangqianjin on 2017/8/2.
 */
public class ExcelUtils {
    public static String detailedDesign="E:\\naviMapII\\17Sum(17Q3)\\Sprint2\\03详细设计和代码\\00详细设计\\detaildesign1.doc";//详细设计路径
    public static String sbl="E:\\naviMapII\\17Sum(17Q3)\\Sprint2\\00项目管理\\00项目计划\\SBL_RD_T110-SP2.xls";
    public static String unitTestPlan="E:\\naviMapII\\17Sum(17Q3)\\Sprint2\\04单元验证\\00测试计划\\计划.doc";
    public static String unitTestCase="";
    public static String unitTestResult="";
    private static Map<String,String>kvfunctionCase;//函数和用例号
    private static Map<String,String>kvucHuman;//UC和开发人员
    private static Map<String,String>kvfunctionTestHuman;//函数和测试人员
    public static void main(String[]args) throws Exception {
        docHelper dh=new docHelper();
        //docHelper.readwordTable(detailedDesign);
        //函数与UC号
        kvfunctionCase=docHelper.readWordCell(detailedDesign);
        int kk=0;
        //UC号与开发人员
        kvucHuman=excelHelper.SBLlist(sbl);
        //读取单元测试计划列表,
        TableIterator tb=docHelper.getdocTable(unitTestPlan);
        // 函数
        List<String>mv=dh.computeContent(tb);
        //请求分配开发人员和测试人员
        List<fncodeuserdevtest>diccolumn=new ArrayList<fncodeuserdevtest>();
        for(String f:kvfunctionCase.keySet()){//函数号与uc号
            fncodeuserdevtest fd=new fncodeuserdevtest();
            fd.setFunctionCode(f);//函数号
            for(String uc :kvucHuman.keySet()){//uc号与开发人员
                String uc2=kvfunctionCase.get(f);
                String hm=kvucHuman.get(uc);
                if(uc.equals(uc2)){
                    fd.setDever(kvucHuman.get(uc));
                    fd.setUsercode(uc);
                    if(hm.equals("武光耀")){
                        fd.setTester("王前进");
                    }
                    if(hm.equals("王前进")){
                        fd.setTester("武光耀");
                    }

                }
            }
            diccolumn.add(fd);
        }
        List<fncodeuserdevtest>result=new ArrayList<fncodeuserdevtest>();
        for(String temp:mv){
            fncodeuserdevtest tp=new fncodeuserdevtest();
            for(fncodeuserdevtest mt:diccolumn){
                if(temp.equals(mt.getFunctionCode())){
                    tp.setFunctionCode(temp);
                    tp.setDever(mt.getDever());
                    tp.setTester(mt.getTester());
                    tp.setUsercode(mt.getUsercode());
                   System.out.println(mt.getUsercode()+"\t"+mt.getFunctionCode()+"\t"+mt.getDever()+"\t"+mt.getTester());

                }
            }
        }
//        for(fncodeuserdevtest mt:diccolumn){
//            System.out.println(mt.getUsercode()+"\t"+mt.getFunctionCode()+"\t"+mt.getDever()+"\t"+mt.getTester());
//        }
    }
}
