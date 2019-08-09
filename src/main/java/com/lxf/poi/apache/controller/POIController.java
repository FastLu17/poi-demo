package com.lxf.poi.apache.controller;

import com.lxf.poi.util.POIUtils;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.*;
import org.apache.poi.xwpf.usermodel.*;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.*;
import java.util.*;

/**
 * @author 小66
 * @Description
 * @create 2019-08-08 10:43
 **/
@RestController
public class POIController {

    /**
     * 获取.doc文档中的所有表格的数据、
     *
     * @return 表格数据
     * @throws Exception
     */
    @GetMapping("tables")
    public List<LinkedHashMap<String, Object>> getTables() throws Exception {
        File file = new File("C:\\Users\\Administrator\\Desktop\\POI\\HWPF测试写入.doc");
        FileInputStream inputStream = new FileInputStream(file);
        HWPFDocument document = new HWPFDocument(inputStream);
        Range range = document.getRange();
        TableIterator tableIterator = new TableIterator(range);

        List<LinkedHashMap<String, Object>> tableList = new ArrayList<>();
        while (tableIterator.hasNext()) {
            //使用: row-column 作为 key、(需要有序,使用LinkedHashMap)
            LinkedHashMap<String, Object> tableMap = new LinkedHashMap<>();
            Table table = tableIterator.next();
            for (int i = 0; i < table.numRows(); i++) {
                TableRow row = table.getRow(i);//TableRow表示:表格的一整行.
                for (int cell = 0; cell < row.numCells(); cell++) {
                    TableCell tableCell = row.getCell(cell);
                    // tableCell.text(): 表示单元格内容、
                    tableMap.put(i + "-" + cell, tableCell.text().trim());
                    // tableCell.numParagraphs()==1、因此不需要再进行遍历、
//                    System.out.println("tableCell.numParagraphs() = " + tableCell.numParagraphs());
//                    for (int para = 0; para < tableCell.numParagraphs(); para++) {
//                        Paragraph paragraph = tableCell.getParagraph(para);
//                    }
                }
            }
            tableList.add(tableMap);
        }
        System.out.println("tableList = " + tableList);

        document.close();
        inputStream.close();
        return tableList;
    }

    /**
     * 在.doc文档中添加表格、
     * <p>
     * new HWPFDocument(inputStream)-->不能打开空的文件。
     *
     * @return
     */
    @GetMapping("/create")
    public String createDOC() throws Exception {
        //新建空白文档、
        File file = new File("C:\\Users\\Administrator\\Desktop\\POI\\HWPF测试写入.doc");
        if (!file.exists()) {
            boolean newFile = file.createNewFile();
            if (!newFile)
                throw new RuntimeException("文件不存在,但是创建文件失败");
        }
        FileInputStream inputStream = new FileInputStream(file);
        HWPFDocument document = new HWPFDocument(inputStream);
        Range range = document.getRange();

        List<Map<String, Object>> mapList = new ArrayList<>();
        Map<String, Object> tableHead = new HashMap<>();
        tableHead.put("name", "姓名");
        tableHead.put("age", "年龄");
        tableHead.put("address", "地址");
        Map<String, Object> userMap = new HashMap<>();
        userMap.put("name", "Jack");
        userMap.put("age", 18);
        userMap.put("address", "北京");
        Map<String, Object> userMap2 = new HashMap<>();
        userMap2.put("name", "Lucy");
        userMap2.put("age", 23);
        userMap2.put("address", "重庆");

        mapList.add(tableHead);
        mapList.add(userMap);
        mapList.add(userMap2);

        int column = 3;
        int row = 4;
        Table table = range.insertTableBefore((short) column, row);
        for (int i = 0; i < row; i++) {
            TableRow tableRow = table.getRow(i);//获取 行
            for (int j = 0; j < column; j++) {
                TableCell cell = tableRow.getCell(j);//获取 单元格、
//                Map<String, Object> map = mapList.get(i);
//                Set<String> keys = map.keySet();
//                cell.replaceText("AAA", true);
                cell.getParagraph(0).getCharacterRun(0).insertBefore("AAAA");
            }
        }
        FileOutputStream outputStream = new FileOutputStream(file);
        document.write(outputStream);
        document.close();
        outputStream.close();
        inputStream.close();
        return "";
    }

    /**
     * 写入.doc格式文件： 使用HWPFDocument
     *
     * @return
     * @throws Exception
     */
    @GetMapping("/write")
    public String writeData() throws Exception {
        File file = new File("C:\\Users\\Administrator\\Desktop\\POI\\HWPF测试写入.doc");
        FileInputStream inputStream = new FileInputStream(file);
        HWPFDocument document = new HWPFDocument(inputStream);

        String doc = document.getText().toString().trim().replaceAll("\u0007", ",");
        System.out.println("doc = " + doc);

        List<Map<String, Object>> mapList = new ArrayList<>();
        Map<String, Object> userMap = new HashMap<>();
        userMap.put("name", "Jack");
        userMap.put("age", 18);
        userMap.put("address", "北京");
        Map<String, Object> userMap2 = new HashMap<>();
        userMap2.put("name", "Lucy");
        userMap2.put("age", 23);
        userMap2.put("address", "重庆");
        mapList.add(userMap);
        mapList.add(userMap2);
        Range range = document.getRange();

        //获取段落对象、
        Paragraph paragraph = range.getParagraph(0);
        /*
         *   使用TableIterator tableIterator = new TableIterator(range);来获取Table对象、
         * */
        Table table = range.getTable(paragraph);//从段落中获取表、
        int numRows = table.numRows();
        System.out.println("numRows = " + numRows);
        TableRow row = table.getRow(0);//获取到表格的某一行、
        System.out.println("row = " + row);

        //将List中的数据填充到表格中去、(替换每行的占位符)
        for (int i = 1; i < numRows; i++) {
            TableRow tableRow = table.getRow(i);
            Map<String, Object> map = mapList.get(i - 1);
            for (String key : map.keySet()) {
                String placeHolder = "${" + key + "}";
                tableRow.replaceText(placeHolder, map.get(key).toString());
            }
        }
        /*
         *   文件中多个同名的placeHolder会被替换同样的数据。
         *       word文档(表格)中:存在两个人的信息(占位符)、替换后,全都变为Jack,18,北京
         * */
        //一次性对整个Range对象范围内的所有内容同时替换掉、
//        for (Map<String, Object> map : mapList) {
//            for (String key : map.keySet()) {
//                String placeHolder = "${"+key+"}";
//                System.out.println("placeHolder = " + placeHolder);
//                range.replaceText(placeHolder,map.get(key).toString());
//            }
//        }

        //把doc输出到输出流中
        FileOutputStream outputStream = new FileOutputStream(file);
        document.write(outputStream);
        document.close();
        outputStream.close();
        inputStream.close();

        return "已完成内容填充";
    }

    /**
     * 读取.doc格式文件： 使用HWPFDocument
     *
     * @return
     * @throws Exception
     */
    @GetMapping("/read")
    public String readData() throws Exception {
        File file = new File("C:\\Users\\Administrator\\Desktop\\POI\\HWPF测试读取.doc");
        FileInputStream istream = new FileInputStream(file);
        HWPFDocument document = new HWPFDocument(istream);

//        HeaderStories headerStories = new HeaderStories(document);//可以操作页眉和页脚
//        Range footerSubrange = headerStories.getEvenFooterSubrange();
//        headerStories.getHeader(1);//获取第几页的页眉
//        headerStories.getFooter(1);

        Range range = document.getRange();//不含页眉和页脚

        //读取整个word内容(有表格的时候)会出现部分乱码的情况、
        //String doc = document.getText().toString().trim().replaceAll("\u0007",",");
        StringBuilder text = new StringBuilder();
        text.append(range.text());
        System.out.println("text = " + text);

        document.close();
        istream.close();
        return text.toString();
    }

    /**
     * 创建.docx文档、
     *
     * @return
     * @throws Exception
     */
    @GetMapping("/createDOCX")
    public String createDOCX() throws Exception {
        XWPFDocument doc = new XWPFDocument();
        XWPFParagraph p1 = doc.createParagraph();
        p1.setAlignment(ParagraphAlignment.CENTER);
        p1.setBorderBottom(Borders.DOUBLE);
        p1.setBorderTop(Borders.DOUBLE);

        p1.setBorderRight(Borders.DOUBLE);
        p1.setBorderLeft(Borders.DOUBLE);
        p1.setBorderBetween(Borders.SINGLE);

        p1.setVerticalAlignment(TextAlignment.TOP);

        XWPFRun r1 = p1.createRun();
        r1.setBold(true);
        r1.setText("The quick brown fox");
        r1.setBold(true);
        r1.setFontFamily("Courier");
        r1.setUnderline(UnderlinePatterns.DOT_DOT_DASH);
        r1.setTextPosition(100);

        r1.addTab();

        FileOutputStream outputStream = new FileOutputStream("C:\\Users\\Administrator\\Desktop\\POI\\XWPF测试读取.docx");
        doc.write(outputStream);

        doc.close();
        outputStream.close();
        return "";
    }

    /**
     * 填充docx文件模板${}的数据
     *
     * @return
     * @throws Exception
     */
    @GetMapping("/updateDOCX")
    public String updateDOCX() throws Exception {
        Map<String, Object> params = new HashMap<>();
        params.put("name", "Jack");
        params.put("age", 18);
        String filePath = "C:\\Users\\Administrator\\Desktop\\POI\\XWPF测试更新.docx";
        InputStream is = new FileInputStream(filePath);
        XWPFDocument doc = new XWPFDocument(is);
        POIUtils poiUtils = new POIUtils();
        //替换段落里面的变量
        poiUtils.resetParagraphDOCX(doc, params);
        //替换表格里面的变量
        poiUtils.resetTableDOCX(doc, params);
        OutputStream os = new FileOutputStream(filePath);

        doc.write(os);
        os.close();
        is.close();
        return "";
    }

    @GetMapping("/insert")
    public String insertData() throws Exception {
        POIUtils poiUtils = new POIUtils();
        List<String> tableHeader = new ArrayList<>();
        tableHeader.add("name");
        tableHeader.add("age");
        tableHeader.add("address");

        List<Map<String, Object>> mapList = new ArrayList<>();//数据库查询获得、
        Map<String, Object> params = new HashMap<>();
        params.put("name", "Jack");
        params.put("age", 18);
        params.put("address", "重庆");
        Map<String, Object> params2 = new HashMap<>();
        params2.put("name", "Lucy");
        params2.put("age", 21);
        params2.put("address", "北京");
        mapList.add(params);
        mapList.add(params2);

        poiUtils.dataInsertIntoTable("XWPF测试新建表格","用户信息表", tableHeader, mapList);
        return "";
    }
}
