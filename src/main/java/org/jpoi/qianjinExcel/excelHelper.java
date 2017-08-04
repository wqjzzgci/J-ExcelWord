package org.jpoi.qianjinExcel;

import com.alibaba.fastjson.JSON;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.util.*;

/**
 * Created by wangqianjin on 2017/7/12.
 */
public class excelHelper {
    protected Logger log = LoggerFactory.getLogger(getClass());
    List<String> m_unitTestCode=new ArrayList<String>();
    List<UserCaseCode>mucode=new ArrayList<UserCaseCode>();
    List<DataInfo>mdatainfo=new ArrayList<DataInfo>();
    public static void main(String[]args) throws IOException, InvalidFormatException {
        List<String>fileList=new ArrayList<String>();
        fileList.add("D:\\吴文斌2\\17WINSP1\\UTRc_RD_T117-SP1.xls");
        fileList.add("D:\\吴文斌2\\17SUMSP2\\UTRc_RD_T110-SP2.xls");
        fileList.add("D:\\吴文斌2\\17SUMSP1\\UTRc_RD_T110-SP1.xls");
        fileList.add("D:\\吴文斌2\\17SPRSP4\\UTRc_RD_T102-SP4.xls");
        fileList.add("D:\\吴文斌2\\17SPRSP3\\UTRc_RD_T102-SP3.xls");
        fileList.add("D:\\吴文斌2\\16WINSP2\\UTRc_RD_T102-SP2.xls");
        fileList.add("D:\\吴文斌2\\16WINSP1\\UTRc_RD_T102-SP1.xls");
        new excelHelper().ReadExcel(fileList);
    }
    private void ReadExcel(List<String>fList)throws IOException, InvalidFormatException{

        //Map<String,String>unitcaseValue=new Hashtable<String, String>();
        for(String mp:fList) {
            InputStream inputStream = new FileInputStream(mp);
            Workbook workbook = WorkbookFactory.create(inputStream);
            Sheet sheet = workbook.getSheetAt(3);
            DataFormatter formatter = new DataFormatter();
            for (Row row : sheet) {
                Cell cell = row.getCell(2);
                if (cell!=null&&cell.getCellType() == Cell.CELL_TYPE_STRING) {
                    String tmp = cell.getRichStringCellValue().getString();
                    if (tmp.startsWith("S")) {
                        String bc="SD015016001001";
                        if(tmp.equals(bc)){
                            int kka=0;
                        }
                        if(m_unitTestCode.indexOf(tmp)==-1){
                            m_unitTestCode.add(tmp);
                        }
                    } else {
                        log.info("未加入的数据是" +mp);
                    }
                } else {
                    log.info("未加入的数据是" +mp);
                }
            }
            workbook.close();
            workbook=null;
            inputStream.close();
            inputStream=null;
            sheet=null;
        }
            List<String>useFiles=new ArrayList<String>();
            //useFiles.add("D:\\吴文斌2\\17WINSP1\\函数用例编号副本.xls");
            //useFiles.add("D:\\吴文斌2\\17SUMSP2\\函数用例编号副本.xls");
           //useFiles.add("D:\\吴文斌2\\17SUMSP1\\函数用例编号副本.xls");
            useFiles.add("D:\\吴文斌2\\17SPRSP4\\函数用例编号副本.xls");
            //useFiles.add("D:\\吴文斌2\\17SPRSP3\\函数用例编号副本.xls");
           // useFiles.add("D:\\吴文斌2\\16WINSP2\\函数用例编号副本.xls");
          // useFiles.add("D:\\吴文斌2\\16WINSP1\\函数用例编号副本.xls");
            for(String userfile:useFiles){
            InputStream userinputStream = new FileInputStream(userfile);
            Workbook userworkbook = WorkbookFactory.create(userinputStream);
            Sheet usersheet = userworkbook.getSheetAt(0);
                int errorLine=0;
            for (Row row : usersheet) {
                try {
                    UserCaseCode tmpuc = new UserCaseCode();
                    errorLine++;
                    Cell cell = row.getCell(0);
                    if(cell!=null) {
                        tmpuc.setFunCode(cell.getRichStringCellValue().getString());
                    }
                    else {
                        continue;
                    }
                    Cell cell1 = row.getCell(1);
                    if(cell1!=null) {
                        tmpuc.setUserCaseCode(cell1.getRichStringCellValue().getString());
                    }
                    else {
                        continue;
                    }
                    Cell cell2 = row.getCell(2);
                    if(cell2!=null) {
                        String tmversion = cell2.getRichStringCellValue().getString();
                        tmpuc.setDaversion(tmversion);
                    }
                    else {
                        continue;
                    }
                    mucode.add(tmpuc);
                } catch (Exception ex) {
                    log.info("错误" + row.getCell(0).getStringCellValue());
                }

            }
                userinputStream.close();
                userworkbook.close();
        }
      // CreateDataInfo();
       CreateExcel1();
    }
    private void CreateDataInfo(){
        Collections.sort(m_unitTestCode);
        for(String str:m_unitTestCode){
            Map<String,Integer>phValue=new HashMap<String,Integer>();
            for(UserCaseCode utcode:mucode){
                if(str.substring(0,11).equalsIgnoreCase(utcode.getFunCode().toString())){
                    DataInfo dt=new DataInfo();
                    dt.setClassCode(str.substring(0,8));//类编号
                    dt.setFunctionCode(str.substring(0, 11));
                    dt.setUnitCode(str);
                    dt.setUsercaseCode(utcode.getUserCaseCode());
                    dt.setDaVersion(utcode.getDaversion());
                    if(utcode.getUserCaseCode().equals("RELATIONSHIP_BUILDING_POI_UC1022")){
                        int kkk=1;
                    }
                    if(utcode.getDaversion().equalsIgnoreCase(DataVersion.YILIUWINSP1)){
                        phValue.put(utcode.getUserCaseCode(),ColumNum.colNum16WINSP1);
                    }
                    else if(utcode.getDaversion().equalsIgnoreCase(DataVersion.YIQIWINSP1)){
                        phValue.put(utcode.getUserCaseCode(),ColumNum.colNum17WINSP1);
                    }
                    else if(utcode.getDaversion().equalsIgnoreCase(DataVersion.YILIUWINSP2)){
                        phValue.put(utcode.getUserCaseCode(),ColumNum.colNum16WINSP2);
                    }
                    if(utcode.getDaversion().equalsIgnoreCase(DataVersion.YIQISPRSP3)){
                        phValue.put(utcode.getUserCaseCode(),ColumNum.colNum17SPRSP3);
                    }
                    else if(utcode.getDaversion().equalsIgnoreCase(DataVersion.YIQISPRSP4)){
                        phValue.put(utcode.getUserCaseCode(),ColumNum.colNum17SPRSP4);
                    }
                    else if(utcode.getDaversion().equalsIgnoreCase(DataVersion.YIQISUMSP1)){
                        phValue.put(utcode.getUserCaseCode(),ColumNum.colNum17SUMSP1);
                    }
                    else if(utcode.getDaversion().equalsIgnoreCase(DataVersion.YIQISUMSP2)){
                        phValue.put(utcode.getUserCaseCode(),ColumNum.colNum17SUMSP2);
                    }
                    dt.setUcollumn(phValue);
                    mdatainfo.add(dt);
                }
            }
        }
    }
    public void CreateExcel() throws IOException {

        // 创建
        HSSFWorkbook wb = new HSSFWorkbook();
        HSSFSheet sheet = wb.createSheet();
        // 创建单元格样式
        HSSFCellStyle titleCellStyle = wb.createCellStyle();
        // 指定单元格居中对齐，边框为细
        titleCellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        titleCellStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
        titleCellStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
        titleCellStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        titleCellStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
        titleCellStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        // 设置填充色
        titleCellStyle.setFillForegroundColor(HSSFColor.SKY_BLUE.index);
        titleCellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        // 指定当单元格内容显示不下时自动换行
        titleCellStyle.setWrapText(true);
        // 设置单元格字体
        HSSFFont titleFont = wb.createFont();
        titleFont.setFontHeightInPoints((short) 12);
        titleFont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
        titleCellStyle.setFont(titleFont);
        HSSFRow headerRow = sheet.createRow(0);
        HSSFCell headerCell = null;
        String[] titles = { "序号", "模块号","类编号","函数编号","单元测试用例号","17WINSP1","17SUMSP2","17SUMSP1","17SPRSP4","17SPRSP3","16WINSP2","16WINSP1" };
        for (int c = 0; c < titles.length; c++) {
            headerCell = headerRow.createCell(c);
            headerCell.setCellStyle(titleCellStyle);
            headerCell.setCellValue(titles[c]);
            sheet.setColumnWidth(c, (30 * 160));
        }
        // ------------------------------------------------------------------
        // 创建单元格样式
        HSSFCellStyle cellStyle = wb.createCellStyle();
        // 指定单元格居中对齐，边框为细
        cellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        cellStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
        cellStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        // 设置单元格字体
        HSSFFont font = wb.createFont();
        //titleFont.setFontHeightInPoints((short) 11);
        font.setBoldweight(HSSFFont.BOLDWEIGHT_NORMAL);
        cellStyle.setFont(font);
        //List<DataInfo> list = JSON.parseArray(infoStr, DataInfo.class);
        int i=0;
        try {
            for (int r = 0; r < mdatainfo.size(); r++) {
                DataInfo item = mdatainfo.get(r);
                if(item.getClassCode()==null)
                    continue;
                HSSFRow row = sheet.createRow(r + 1);
                HSSFCell cell = null;
                int c = 1;
                //模块编号
                cell = row.createCell(c++);
                cell.setCellStyle(cellStyle);
                cell.setCellValue(item.getClassCode().substring(0, 5));
                //类编号
                cell = row.createCell(c++);
                cell.setCellStyle(cellStyle);
                cell.setCellValue(item.getClassCode());
                //函数编号
                cell = row.createCell(c++);
                cell.setCellStyle(cellStyle);
                cell.setCellValue(item.getFunctionCode());
                //单元测试用例编号
                cell = row.createCell(c++);
                cell.setCellStyle(cellStyle);
                cell.setCellValue(item.getUnitCode());
                if(item.getUsercaseCode().equals("SHOW_AU_MARK_UC0415")){
                    int zbc=0;
                }
                //用例编号
                int kv=0;
                for (String key : item.getUcollumn().keySet()) {
                    int num=item.getUcollumn().get(key) + 4+kv;
                    if(key.equals("SHOW_AU_MARK_UC0415")){
                        int  kbby=2;
                    }
                    cell = row.createCell(num);
                    cell.setCellStyle(cellStyle);
                    cell.setCellValue(key);
                    kv++;
                }

            }

            FileOutputStream fileOut = new FileOutputStream("E:/test/test1.xls");
            wb.write(fileOut);
            fileOut.close();
            System.out.println("Done");
        }
        catch (Exception ex){
            System.out.println(i+"行");
        }
    }
    public void CreateExcel1() throws IOException {

        Collections.sort(m_unitTestCode);
        // 创建
        HSSFWorkbook wb = new HSSFWorkbook();
        HSSFSheet sheet = wb.createSheet();
        // 创建单元格样式
        HSSFCellStyle titleCellStyle = wb.createCellStyle();
        // 指定单元格居中对齐，边框为细
        titleCellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        titleCellStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
        titleCellStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
        titleCellStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        titleCellStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
        titleCellStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        // 设置填充色
        titleCellStyle.setFillForegroundColor(HSSFColor.SKY_BLUE.index);
        titleCellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        // 指定当单元格内容显示不下时自动换行
        titleCellStyle.setWrapText(true);
        // 设置单元格字体
        HSSFFont titleFont = wb.createFont();
        titleFont.setFontHeightInPoints((short) 12);
        titleFont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
        titleCellStyle.setFont(titleFont);
        HSSFRow headerRow = sheet.createRow(0);
        HSSFCell headerCell = null;
        String[] titles = { "序号", "模块号","类编号","函数编号","单元测试用例号","17WINSP1","17SUMSP2","17SUMSP1","17SPRSP4","17SPRSP3","16WINSP2","16WINSP1" };
        for (int c = 0; c < titles.length; c++) {
            headerCell = headerRow.createCell(c);
            headerCell.setCellStyle(titleCellStyle);
            headerCell.setCellValue(titles[c]);
            sheet.setColumnWidth(c, (30 * 160));
        }
        // ------------------------------------------------------------------
        // 创建单元格样式
        HSSFCellStyle cellStyle = wb.createCellStyle();
        // 指定单元格居中对齐，边框为细
        cellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        cellStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
        cellStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        // 设置单元格字体
        HSSFFont font = wb.createFont();
        //titleFont.setFontHeightInPoints((short) 11);
        font.setBoldweight(HSSFFont.BOLDWEIGHT_NORMAL);
        cellStyle.setFont(font);
        //List<DataInfo> list = JSON.parseArray(infoStr, DataInfo.class);
        int i=0;
        try {
            for(String str:m_unitTestCode){
                HSSFRow row = sheet.createRow(i + 1);
                HSSFCell cell = null;
                int c = 1;
                //模块编号
                cell = row.createCell(c++);
                cell.setCellStyle(cellStyle);
                cell.setCellValue(str.substring(0, 5));
                //类编号
                cell = row.createCell(c++);
                cell.setCellStyle(cellStyle);
                cell.setCellValue(str.substring(0,8));
                //函数编号
                cell = row.createCell(c++);
                cell.setCellStyle(cellStyle);
                cell.setCellValue(str.substring(0, 11));
                //单元测试用例编号
                cell = row.createCell(c++);
                cell.setCellStyle(cellStyle);
                cell.setCellValue(str);

                for(UserCaseCode utcode:mucode){
                    if(str.substring(0,11).equalsIgnoreCase(utcode.getFunCode().toString())){
                        int num=11;
                        cell = row.createCell(num);
                        cell.setCellStyle(cellStyle);
                        cell.setCellValue(utcode.getUserCaseCode());
                        //kv++;
                    }
                }
                i++;
            }
            FileOutputStream fileOut = new FileOutputStream("E:/test/test1.xls");
            wb.write(fileOut);
            fileOut.close();
            System.out.println("Done");
        }
        catch (Exception ex){
            System.out.println(i+"行");
        }
    }
    public static Map<String,String> SBLlist(String sblFile) throws IOException, InvalidFormatException {
        Map<String,String>kmp=new HashMap<String, String>();
            InputStream inputStream = new FileInputStream(sblFile);
            Workbook workbook = WorkbookFactory.create(inputStream);
            Sheet sheet = workbook.getSheetAt(2);
            DataFormatter formatter = new DataFormatter();
            for (Row row : sheet) {
                String k="";
                String v="";
                Cell celluc = row.getCell(9);
                if (celluc!=null&&celluc.getCellType() == Cell.CELL_TYPE_STRING) {
                    k = celluc.getRichStringCellValue().getString();
                    k=k.substring(k.length()-6,k.length());
                    if(!k.contains("UC")){
                        continue;
                    }
                }
                if(k.isEmpty())
                    continue;
                Cell celldever = row.getCell(25);
                if (celldever!=null&&celldever.getCellType() == Cell.CELL_TYPE_STRING) {
                    v = celldever.getRichStringCellValue().getString().trim();
                }
                if(v.isEmpty())
                {
                    continue;
                }
                kmp.put(k,v);
//                while (row.cellIterator().hasNext()){
//                    Cell cell1=row.cellIterator().next();
//                    if (cell!=null&&cell.getCellType() == Cell.CELL_TYPE_STRING){
//                        String tmp = cell.getRichStringCellValue().getString();
//                    }
//                }
                }
            workbook.close();
            workbook=null;
            inputStream.close();
            inputStream=null;
            sheet=null;
        return kmp;
    }
}
