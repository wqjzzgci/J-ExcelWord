package org.jpoi.qianjinExcel;
import java.io.*;
import java.io.File;
import java.io.FileInputStream;

import java.io.FileInputStream;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

/**
 * Created by wangqianjin on 2017/8/2.
 */
public class docHelper {
    public static Map<String,String> readwordTable(String file) throws IOException {
        Map<String,String>kv=new HashMap<String, String>();
        FileInputStream fp=new FileInputStream(file);
        XWPFDocument document = new XWPFDocument(fp);
// 获取所有表格
        List<XWPFTable> tables = document.getTables();
        for (XWPFTable table : tables) {
            // 获取表格的行
            List<XWPFTableRow> rows = table.getRows();
            for (XWPFTableRow row : rows) {
                // 获取表格的每个单元格
                List<XWPFTableCell> tableCells = row.getTableCells();
                for (XWPFTableCell cell : tableCells) {
                    // 获取单元格的内容
                    String text = cell.getText();
                    kv.put(text,text);
                }
            }
        }
        return kv;
    }
    public static Map<String,String> readWordCell(String filePath) {
        Map<String,String>kv=new HashMap<String, String>();
        FileInputStream in = null;
        POIFSFileSystem pfs = null;
        List<String> list = new ArrayList<String>();
        try {
            in = new FileInputStream(filePath);// 载入文档
            pfs = new POIFSFileSystem(in);
            HWPFDocument hwpf = new HWPFDocument(pfs);
            Range range = hwpf.getRange();// 得到文档的读取范围
            TableIterator it = new TableIterator(range);
            // 迭代文档中的表格
            if (it.hasNext()) {
                TableRow tr = null;
                TableCell td = null;
                Paragraph para = null;
                String lineString;
                String cellString;
                Table tb = (Table) it.next();
                // 迭代行，从第2行开始
                for (int i = 0; i < tb.numRows(); i++) {
                    tr = tb.getRow(i);
                    lineString = "";
                    String functioncode="";
                    String userCode="";
                    for (int j = 0; j < tr.numCells(); j++) {
                        td = tr.getCell(j);// 取得单元格
                        // 取得单元格的内容
                        for (int k = 0; k < td.numParagraphs(); k++) {
                            para = td.getParagraph(k);
                            cellString = para.text();
                            if (cellString != null
                                    && cellString.compareTo("") != 0) {
                                if(j==0){
                                    functioncode=cellString.trim();
                                }
                                if(j==4){
                                    String usr=cellString.trim();
                                    userCode=usr.substring(usr.length()-6,usr.length());
                                }
                                // 如果不trim，取出的内容后会有一个乱码字符
                                cellString = cellString.trim() + "|";
                            }
                            lineString += cellString;
                        }
                    }
                    // 去除字符串末尾的一个管道符
                    if (lineString != null && lineString.compareTo("") != 0) {
                        lineString = lineString.substring(0, lineString
                                .length() - 1);
                    }
                    kv.put(functioncode,userCode);
                    list.add(lineString);
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (null != in) {
                try {
                    in.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
        return kv;
    }
    public static TableIterator getdocTable(String fp){
        FileInputStream in = null;
        POIFSFileSystem pfs = null;
        List<String> list = new ArrayList<String>();
        TableIterator it;
        try {
            in = new FileInputStream(fp);// 载入文档
            pfs = new POIFSFileSystem(in);
            HWPFDocument hwpf = new HWPFDocument(pfs);
            Range range = hwpf.getRange();// 得到文档的读取范围
            it= new TableIterator(range);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
            return null;
        } catch (IOException e) {
            e.printStackTrace();
            return null;
        }
        return it;
    }
    public List<String> computeContent(TableIterator it){
        List<String>vlist=new ArrayList<String>();
        if(it.hasNext()){
            TableRow tr = null;
            TableCell td = null;
            Paragraph para = null;
            Table tb = (Table) it.next();
            for (int i = 0; i < tb.numRows(); i++) {
                tr = tb.getRow(i);
                String functioncode="";
                String userCode="";
                if(tr.numCells()==0)
                    continue;
                String function=tr.getCell(0).text().trim();
                int kk=0;
                vlist.add(function);
//                for (int j = 0; j < tr.numCells(); j++) {
//                    td = tr.getCell(j);// 取得单元格
//                    td.text();
               }
        }
       return vlist;
    }
    public static void createTable(XWPFTable xTable,XWPFDocument xdoc){
                String bgColor="111111";
                CTTbl ttbl = xTable.getCTTbl();
                CTTblPr tblPr = ttbl.getTblPr() == null ? ttbl.addNewTblPr() : ttbl.getTblPr();
                CTTblWidth tblWidth = tblPr.isSetTblW() ? tblPr.getTblW() : tblPr.addNewTblW();
                tblWidth.setW(new BigInteger("8600"));
                tblWidth.setType(STTblWidth.DXA);
                setCellText(xdoc, getCellHight(xTable, 0, 0), "序号",bgColor,1000);
                setCellText(xdoc, getCellHight(xTable, 0, 1), "阶段",bgColor,3800);
                setCellText(xdoc, getCellHight(xTable, 0, 2), "计划工作任务",bgColor,3800);
                int number=0;
                for(int i=1;i<5;i++){
                        number++;
                        setCellText(xdoc, getCellHight(xTable, number,0), number+"",bgColor,1000);
                        setCellText(xdoc, getCellHight(xTable, number,1), "階段名稱",bgColor,3800);
                        setCellText(xdoc, getCellHight(xTable, number,2), "任務名稱",bgColor,3800);
                    }
            }
    /**
     118.     *
     119.     * @param xDocument
     120.     * @param cell
     121.     * @param text
     122.     * @param bgcolor
     123.     * @param width
     124.     */
        private static void setCellText(XWPFDocument xDocument, XWPFTableCell cell,
                 String text, String bgcolor, int width) {
                CTTc cttc = cell.getCTTc();
                CTTcPr cellPr = cttc.addNewTcPr();
                cellPr.addNewTcW().setW(BigInteger.valueOf(width));
                XWPFParagraph pIO =cell.addParagraph();
                cell.removeParagraph(0);
                XWPFRun rIO = pIO.createRun();
                rIO.setFontFamily("微软雅黑");
                rIO.setColor("000000");
                rIO.setFontSize(12);
                rIO.setText(text);
            }
//设置表格高度
        private static XWPFTableCell getCellHight(XWPFTable xTable,int rowNomber,int cellNumber){
                XWPFTableRow row = null;
                row = xTable.getRow(rowNomber);
                row.setHeight(100);
                XWPFTableCell cell = null;
                cell = row.getCell(cellNumber);
                return cell;
            }


}
