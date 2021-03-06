package com.lxf.poi.controller;

import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import com.lxf.poi.mapper.UserInfoMapper;
import com.lxf.poi.util.POIWordUtil;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.*;
import org.apache.poi.xwpf.usermodel.*;
import org.springframework.beans.factory.annotation.Autowired;
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
public class POIWordController {

    @Autowired
    private POIWordUtil wordUtil;

    @SuppressWarnings("SpringJavaInjectionPointsAutowiringInspection")
    @Autowired
    private UserInfoMapper mapper;

    private final String BASE_DIRECTORY_PATH = "C:\\Users\\Administrator\\Desktop\\POI\\";
    private final String DOC_TEMPLATE_FILE_PATH = BASE_DIRECTORY_PATH + "HWPF测试模板.doc";
    private String DOCX_TEMPLATE_FILE_PATH = BASE_DIRECTORY_PATH + "XWPF测试模板.docx";

    /**
     * 获取.doc文档中的所有表格的数据、
     *
     * @return 表格数据
     * @throws Exception
     */
    @GetMapping("tables")
    public List<LinkedHashMap<String, Object>> getTables() throws Exception {
        return wordUtil.getTablesDataList(BASE_DIRECTORY_PATH + "HWPF测试写入.doc");
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
        File file = new File(BASE_DIRECTORY_PATH + "HWPF测试写入.doc");
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
            if (i == 0) {
                tableRow.setTableHeader(true);
                tableRow.setRowJustification(100);
            }
            for (int j = 0; j < column; j++) {
                TableCell cell = tableRow.getCell(j);//获取 单元格、
                CharacterRun run = cell.getParagraph(0).getCharacterRun(0);
                run.setFontSize(16);
                run.insertBefore("AAAA");
                run.setBold(true);
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
        FileInputStream inputStream = new FileInputStream(DOC_TEMPLATE_FILE_PATH);
        HWPFDocument document = new HWPFDocument(inputStream);

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

        POIWordUtil poiWordUtil = new POIWordUtil();
        poiWordUtil.resetTableDOC(document, userMap2, 1);

        /*
         *   一次性设置多张相同模板的表的${}变量为不同的数据、
         * */
//        Range range = document.getRange();
//        TableIterator tableIterator = new TableIterator(range);//从段落中获取表、
//        while (tableIterator.hasNext()) {
//            Table table = tableIterator.next();
//            //将List中的数据填充到表格中去、(替换每行的占位符)
//            for (int i = 1; i < table.numRows(); i++) {
//                TableRow tableRow = table.getRow(i);
//                Map<String, Object> map = mapList.get(i - 1);
//                for (String key : map.keySet()) {
//                    String placeHolder = "${" + key + "}";
//                    tableRow.replaceText(placeHolder, map.get(key).toString());
//                }
//            }
//        }

        String doc = document.getText().toString().trim().replaceAll("\u0007", ",").replaceAll(",,", ",");
        //把doc输出到输出流中
        FileOutputStream outputStream = new FileOutputStream(BASE_DIRECTORY_PATH + "HWPF测试resetTableDOC.doc");
        document.write(outputStream);
        document.close();
        outputStream.close();
        inputStream.close();

        return doc;
    }

    /**
     * 读取.doc格式文件： 使用HWPFDocument
     *
     * @return
     * @throws Exception
     */
    @GetMapping("/read")
    public String readData() throws Exception {
        File file = new File(BASE_DIRECTORY_PATH + "HWPF测试读取.doc");
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

        FileOutputStream outputStream = new FileOutputStream(BASE_DIRECTORY_PATH + "XWPF测试新建DOCX文件.docx");
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
        String filePath = BASE_DIRECTORY_PATH + "XWPF测试更新.docx";
        InputStream is = new FileInputStream(filePath);
        XWPFDocument doc = new XWPFDocument(is);
        POIWordUtil poiWordUtil = new POIWordUtil();
        //替换段落里面的变量
        poiWordUtil.resetParagraphDOCX(doc, params);
        //替换表格里面的变量
        poiWordUtil.resetTableDOCX(doc, params);
        OutputStream os = new FileOutputStream(filePath);

        doc.write(os);
        os.close();
        is.close();
        return "";
    }

    @GetMapping("/insert")
    public String insertData() throws Exception {
        POIWordUtil poiWordUtil = new POIWordUtil();
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

        return poiWordUtil.createTableByData("XWPF测试新建表格", "用户信息表",
                tableHeader, mapList, "Consolas", 8000, 18, 14);
    }

    @GetMapping("/resetAllDOC")
    public String resetAllDOC() throws Exception {
        File file = new File(DOC_TEMPLATE_FILE_PATH);
        FileInputStream inputStream = new FileInputStream(file);
        HWPFDocument document = new HWPFDocument(inputStream);

        Map<String, Object> userMap = new HashMap<>();
        userMap.put("name", "Jack");
        userMap.put("age", 18);
        userMap.put("address", "北京");

        POIWordUtil poiWordUtil = new POIWordUtil();
        poiWordUtil.resetAllDOC(document, userMap);//填充文件段落中的占位符、

        String text = document.getRange().text();

        FileOutputStream outputStream = new FileOutputStream(
                BASE_DIRECTORY_PATH + "HWPF测试restParagraphDOC.doc");
        document.write(outputStream);
        poiWordUtil.closeStream(document, outputStream, inputStream);
        return text;
    }

    @GetMapping("/getTablesDataList")
    public String getTablesDataList() throws Exception {
        POIWordUtil poiWordUtil = new POIWordUtil();
        List<LinkedHashMap<String, Object>> tablesDataList = poiWordUtil.getTablesDataList(DOCX_TEMPLATE_FILE_PATH);
        return tablesDataList.toString();
    }

    @GetMapping("/insertNewEmptyRows")
    public String insertNewEmptyRows() throws Exception {
        POIWordUtil poiWordUtil = new POIWordUtil();
        return poiWordUtil.addEmptyRows(DOCX_TEMPLATE_FILE_PATH, 3);
    }

    @GetMapping("/addNotEmptyRows")
    public String addNotEmptyRows() throws Exception {
        List<String> tableHeader = new ArrayList<>();
        tableHeader.add("name");
        tableHeader.add("age");
        tableHeader.add("address");
        /*
         *   Map包含name,age,address三个Key、
         * */
        List<Map<String, Object>> mapList = mapper.selectAllResultMap();//4.6W数据

        //同步调用,海量数据-->效率低下,还可能出现OutOfMemoryError异常、
        wordUtil.addNotEmptyRows(DOCX_TEMPLATE_FILE_PATH, tableHeader, mapList);
        return "已导出" + mapList.size() + 1 + "行数据到表格中";
    }

    @GetMapping("/asyncAddNotEmptyRows")
    public String asyncAddNotEmptyRows() throws Exception {
        List<String> tableHeader = new ArrayList<>();
        tableHeader.add("name");
        tableHeader.add("age");
        tableHeader.add("address");
        /*
         *   Map包含name,age,address三个Key、
         * */
        List<Map<String, Object>> mapList = mapper.selectAllResultMap();//4.6W数据

        //异步调用、
        wordUtil.asyncAddNotEmptyRows(DOCX_TEMPLATE_FILE_PATH, tableHeader, mapList);
        return "已获得" + mapList.size() + 1 + "行数据,正在导出到表格中";
    }

    @GetMapping("/jsonData")
    public String getJsonData() {
//        String strJson = "{\"a\":\"1\",\"b\":\"2\",\"c\":\"3\",\"d\":[{\"a1\":\"  aa  aa  \",\"a2\":\"  a2a  a2a  \",\"a3\":\"  a33a  a33a  33\"},{\"a1\":\"  aa  aa  \",\"a2\":\"  a2a  a2a  \",\"a3\":\"  a33a  a33a  33\"}]}";
        String strJson = "{\"aa\":[\"str1\",\"  str2  str2  \"],\"d\":[[{\"aa\":[{\"map\":\"  map  map  \"}]},{\"bb\":\"  bb  bb  \"}],[{\"a2a\":\"  a2a  a2a  \"},{\"b2b\":\"  b2b  b2b  \"}]]}";
        Map<String, Object> params = new HashMap<>();
        params.put("a1", "aaa  aaa  ");
        params.put("a2", "  a222aa  aaa  ");
        params.put("a3", "   a333aa  aaa  ");
        Map<String, Object> map = new HashMap<>();
        map.put("b1", "bb bb  ");
        map.put("b2", "  b222b bb  ");
        map.put("b3", "    b3333b bb  ");
        JSONArray jsonArray = new JSONArray();
        jsonArray.add(map);
        params.put("list", jsonArray);

        String mapString = JSONObject.toJSONString(params);

//        JSONObject jsonObjectMap = new JSONObject(params); //这种方式生成的JSONObject无法正常解析、
//        JSONObject jsonObject = JSON.parseObject(mapString);
        JSONObject jsonObject = JSONObject.parseObject(strJson);
        System.out.println("jsonObject = " + jsonObject);
        trimJsonObject(jsonObject);
        System.out.println("jsonObject After = " + jsonObject);
        return "AA";
    }

    /**
     * 利用fastJson解析请求JSON数据,去除请求参数两端的空格、
     *
     * @param params JSON数据、
     */
    private void trimJsonObject(Map<String, Object> params) {
        for (Map.Entry<String, Object> entry : params.entrySet()) {
            if (entry.getValue() instanceof String) {
                params.put(entry.getKey(), entry.getValue().toString().trim());
                continue;
            }
            //JSONObject是Map的实现类、但不是Map,如果value是Map则无法解析、
            if (entry.getValue() instanceof JSONObject) {
                trimJsonObject((JSONObject) entry.getValue());
                continue;
            }
            //JSONArray是List的实现类、但不是List,如果value是List则无法解析、
            if (entry.getValue() instanceof JSONArray) {
                trimJsonArray((JSONArray) entry.getValue());
            }
        }
    }

    /**
     * List覆盖当前index的值、需要使用fori的循环方式进行、
     *      否则抛出:java.util.ConcurrentModificationException: null
     * @param jsonArray
     */
    private void trimJsonArray(JSONArray jsonArray) {
        for (int i = 0; i < jsonArray.size(); i++) {
            if (jsonArray.get(i) instanceof String) {
                Object remove = jsonArray.remove(i);//返回被删除的对象、
                jsonArray.add(i, remove.toString().trim());
                continue;
            }
            if (jsonArray.get(i) instanceof JSONObject) {
                trimJsonObject((JSONObject) jsonArray.get(i));
                continue;
            }
            if (jsonArray.get(i) instanceof JSONArray) {
                trimJsonArray((JSONArray) jsonArray.get(i));
            }
        }
        /*
         * TODO:这样循环时处理remove()和add()-->抛出异常:java.util.ConcurrentModificationException: null
         * */
//        for (Object obj : jsonArray) {
//            if (obj instanceof String) {
//                int index = jsonArray.indexOf(obj);
//                boolean remove = jsonArray.remove(obj);//返回boolean值、
//                jsonArray.add(index, obj.toString().trim());
//                continue;
//            }
//            if (obj instanceof JSONObject) {
//                trimJsonObject((JSONObject) obj);
//                continue;
//            }
//            if (obj instanceof JSONArray) {
//                trimJsonArray((JSONArray) obj);
//            }
//        }
    }
}
