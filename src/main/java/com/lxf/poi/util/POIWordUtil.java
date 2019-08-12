package com.lxf.poi.util;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.*;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.springframework.scheduling.annotation.Async;
import org.springframework.stereotype.Component;
import org.springframework.util.StringUtils;

import java.io.*;
import java.math.BigInteger;
import java.util.*;
import java.util.concurrent.Executors;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;

/**
 * 按需将RunTimeException替换为自定义异常、
 * <p>
 * 按需添加异步的方法、类似asyncAddNotEmptyRows()
 * <p>
 * HWPFDocument修改文本信息：使用(Range、Paragraph、CharacterRun中的)insertBefore()和insertAfter()方法可以修改相应文本信息。
 * <p>
 * XWPFDocument修改文本信息：从XWPFParagraph中，可以获取组成文本信息的XWPFRun要素。
 *                          如果要添加新文本，调用createRun()方法或者文本末端添加一个XWPFRun要素，
 *                          insertNewRun(int)可以paragraph的指定位置添加一个XWPFRun要素。一旦有了XWPFRun，
 *                          可以调用其setText(String)方法修改文本内容，如果想增加一个空白要素，如tabs或则line breaks，
 *                          需要调用addTab()和addCarriageReturn()方法。
 * @author 小66
 * @Description
 * @create 2019-08-09 15:17
 **/
@Component
@Slf4j
public class POIWordUtil {

    private final String BASE_URL = "C:\\Users\\Administrator\\Desktop\\POI\\";

    /**
     * 异步调用：常用于无返回值的方法、
     * 使用Future对象用作异步返回值,如果执行了get()则会一直阻塞,无法进行异步调用、
     */
    @Async("taskExecutor")
    public void asyncAddNotEmptyRows(String absPath, List<String> tableHeader,
                                     List<Map<String, Object>> params) throws IOException {
        String filePath = addNotEmptyRows(absPath, tableHeader, params);
        //需要通过WebSocket反馈给用户后台异步任务已完成.
    }

    /**
     * 为指定index的表在表的fromRow行后添加add条数据
     *
     * @param absPath     docx文件绝对路径
     * @param add         添加的总行数
     * @param fromRow     添加开始的位置
     * @param tableHeader 表头的名称(第一行的单元格内容)
     *                    Tips:** tableHeader 需要 与 params的key相同。(此处不会更改tableHeader,随意命名即可)
     * @param params      需要填充的数据
     * @param tableIndex  第 index 张表格需要添加数据、从1开始.
     * @return absPath
     * @throws IOException IO
     */
    public String addNotEmptyRows(String absPath, int add, int fromRow, List<String> tableHeader,
                                  List<Map<String, Object>> params, int tableIndex) throws IOException {
        long startTime = System.currentTimeMillis();
        if (StringUtils.isEmpty(absPath)) {
            throw new RuntimeException("文件的绝对路径不能为空");
        }
        if (!absPath.endsWith(".docx") && !absPath.endsWith(".DOCX"))
            throw new RuntimeException("无法处理非docx文件");
        if (tableIndex <= 0)
            throw new RuntimeException("tableIndexArr必须大于等于1");

        FileInputStream inputStream = new FileInputStream(absPath);
        XWPFDocument docx = new XWPFDocument(inputStream);
        List<XWPFTable> tables = docx.getTables();
        if (tableIndex - 1 > tables.size()) {
            throw new RuntimeException("文件中表格数量低于" + tableIndex + "个");
        }

        XWPFTable table = tables.get(tableIndex - 1);
        if (fromRow == 0) //表最后添加数据
            insertNotEmptyRows(table, add, table.getRows().size(), tableHeader, params);
        else
            insertNotEmptyRows(table, add, fromRow, tableHeader, params);

        //正常操作是写入到读取的文件中去、
        //FileOutputStream outputStream = new FileOutputStream(absPath);
        FileOutputStream outputStream = new FileOutputStream(BASE_URL + "XWPF测试insertNotEmptyRows.docx");
        docx.write(outputStream);
        closeStream(docx, outputStream, inputStream);

        long endTime = System.currentTimeMillis();
        log.info("totalTime = {} 秒", (endTime - startTime) / 1000);
        return absPath;
    }

    /**
     * 不传入第几张表时,默认第一张表添加数据、
     *
     * @param absPath     docx文件绝对路径
     * @param add         添加的总行数
     * @param fromRow     添加开始的位置
     * @param tableHeader 表头的名称(第一行的单元格内容)
     * @param params      需要填充的数据
     * @return absPath
     * @throws IOException IO
     */
    public String addNotEmptyRows(String absPath, int add, int fromRow, List<String> tableHeader,
                                  List<Map<String, Object>> params) throws IOException {
        return addNotEmptyRows(absPath, add, fromRow, tableHeader, params, 1);
    }

    /**
     * 为指定index的表在表的最后添加params.size()+1条数据
     *
     * @param absPath     docx文件绝对路径
     * @param tableHeader 表头的名称(第一行的单元格内容)
     *                    Tips:** tableHeader 需要 与 params的key相同。(此处不会更改tableHeader,随意命名即可)
     * @param params      需要填充的数据
     * @param tableIndex  第 index 张表格需要添加数据、从1开始.
     * @return absPath
     * @throws IOException IO
     */
    public String addNotEmptyRows(String absPath, List<String> tableHeader,
                                  List<Map<String, Object>> params, int tableIndex) throws IOException {
        return addNotEmptyRows(absPath, params.size(), 0, tableHeader, params, tableIndex);
    }

    /**
     * 不传入第几张表时,默认第一张表添加数据、
     *
     * @param absPath     docx文件绝对路径
     * @param tableHeader 表头的名称(第一行的单元格内容)
     * @param params      需要填充的数据
     * @return absPath
     * @throws IOException IO
     */
    public String addNotEmptyRows(String absPath, List<String> tableHeader,
                                  List<Map<String, Object>> params) throws IOException {
        return addNotEmptyRows(absPath, params.size(), 0, tableHeader, params, 1);
    }

    /**
     * 默认为第一张表添加同属性的空行、
     *
     * @param absPath    docx文件绝对路径
     * @param addRowsNum 需要添加的行数
     * @return absPath
     * @throws IOException IO
     */
    public String addEmptyRows(String absPath, int addRowsNum) throws IOException {
        return addEmptyRows(absPath, addRowsNum, 1);
    }

    /**
     * 为指定index的表添加同属性的空行、
     *
     * @param absPath       docx文件绝对路径
     * @param addRowsNum    需要添加的行数
     * @param tableIndexArr 第 index 张表格需要添加数据、从1开始.
     * @return absPath
     * @throws IOException IO
     */
    public String addEmptyRows(String absPath, int addRowsNum, Integer... tableIndexArr) throws IOException {
        if (StringUtils.isEmpty(absPath)) {
            throw new RuntimeException("文件的绝对路径不能为空");
        }
        if (!absPath.endsWith(".docx") && !absPath.endsWith(".DOCX"))
            throw new RuntimeException("无法处理非docx文件");
        List<Integer> list = Arrays.stream(tableIndexArr).filter(integer -> integer < 1).collect(Collectors.toList());
        if (list.size() > 0)
            throw new RuntimeException("tableIndexArr必须大于等于1");
        FileInputStream inputStream = new FileInputStream(absPath);
        XWPFDocument docx = new XWPFDocument(inputStream);
        List<XWPFTable> tables = docx.getTables();
        for (int tableIndex : tableIndexArr) {
            if (tableIndex - 1 > tables.size()) {
                throw new RuntimeException("文件中表格数量低于" + tableIndex + "个");
            }
            XWPFTable table = tables.get(tableIndex);
            //新增EmptyRows、
            insertOrRemoveEmptyRows(table, addRowsNum, table.getRows().size());
        }
        //正常操作是写入到读取的文件中去、
        //FileOutputStream outputStream = new FileOutputStream(absPath);
        FileOutputStream outputStream = new FileOutputStream(BASE_URL + "XWPF测试insertEmptyRows.docx");
        docx.write(outputStream);
        closeStream(docx, outputStream, inputStream);
        return absPath;
    }

    /**
     * 创建docx文件并插入一张表格,填充表格内容和数据
     *
     * @param fileName       需要生成的文件名
     * @param tableName      文件中的表名
     * @param tableHeader    表格的表头,Tips:** tableHeader 需要 与 params的key相同。(表格第一行的内容是由此对象填充,按需传递)
     * @param params         表格的内容
     * @param fontFamily     字体样式
     * @param tableWidth     表格宽度,例如:8000
     * @param headerFontSize 字体大小(表头)
     * @param bodyFontSize   字体大小(表体)
     * @return 文件路径
     * @throws IOException IO
     */
    public String createTableByData(String fileName, String tableName, List<String> tableHeader, List<Map<String, Object>> params, String fontFamily, int tableWidth, int headerFontSize, int bodyFontSize) throws IOException {

        String filePath = BASE_URL + fileName + ".docx";

        XWPFDocument docx = new XWPFDocument();
        //创建标题段落
        XWPFParagraph titlePara = docx.createParagraph();
        XWPFRun titleRun = titlePara.createRun();
        titleRun.setText(tableName);
        //利用文本对象(XWPFRun)设置样式、
        titleRun.setFontSize(20);
        titleRun.setBold(true);//字体是否加粗
        titlePara.setAlignment(ParagraphAlignment.CENTER);//段落居中

        XWPFTable table = docx.createTable(params.size() + 1, tableHeader.size());
        //表格属性
        CTTblPr ctTblPr = table.getCTTbl().addNewTblPr();
        //设置表格宽度
        CTTblWidth width = ctTblPr.addNewTblW();
        width.setW(BigInteger.valueOf(tableWidth));
        //设置表格宽带为固定、
        width.setType(STTblWidth.DXA);//STTblWidth.AUTO:自动-->宽度无效.

        //TODO: 如何设置表格位置?

        List<XWPFTableRow> tableRows = table.getRows();

        //获取表头
        XWPFTableRow header = tableRows.get(0);
        List<XWPFTableCell> tableCells = header.getTableCells();
        //设置表头的样式和属性
        for (int i = 0; i < tableHeader.size(); i++) {
            XWPFTableCell headerCell = tableCells.get(i);
            //垂直居中
            headerCell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
            XWPFParagraph headerPara = headerCell.addParagraph();//获取单元格的段落、
            XWPFRun headerRun = headerPara.createRun();//XWPFRun是最小的单位.(文本)
            headerRun.setText(tableHeader.get(i));
            headerRun.setBold(true);
            headerRun.setFontSize(headerFontSize);
            headerRun.setFontFamily(fontFamily);
            headerRun.setColor("66FF66");
            headerRun.setShadow(true);//阴影-->按需设置
            //水平居中
            headerPara.setAlignment(ParagraphAlignment.CENTER);
            //垂直居中
            headerPara.setVerticalAlignment(TextAlignment.CENTER);
        }

        //设置表体的样式和属性
        for (int i = 0; i < params.size(); i++) {
            Map<String, Object> param = params.get(i);
            XWPFTableRow tableRow = tableRows.get(i + 1);
            for (int j = 0; j < tableRow.getTableCells().size(); j++) {
                XWPFTableCell cell = tableRow.getCell(j);
                XWPFParagraph cellParagraph = cell.addParagraph();
                //垂直居中
                cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
                XWPFRun cellRun = cellParagraph.createRun();
                cellRun.setText(param.get(tableHeader.get(j)).toString());
                cellRun.setBold(false);
                cellRun.setFontSize(bodyFontSize);
                cellRun.setFontFamily(fontFamily);
                cellRun.setUnderline(UnderlinePatterns.SINGLE);//下划线-->按需设置

                cellParagraph.setVerticalAlignment(TextAlignment.CENTER);
                cellParagraph.setAlignment(ParagraphAlignment.CENTER);
            }
        }

        FileOutputStream outputStream = new FileOutputStream(filePath);
        docx.write(outputStream);
        closeStream(docx, outputStream);
        return filePath;
    }

    /**
     * 替换docx文件中表格的${}变量
     *
     * @param docx   要替换的docx文件
     * @param params 参数
     */
    public void resetTableDOCX(XWPFDocument docx, Map<String, Object> params) {
        Iterator<XWPFTable> iterator = docx.getTablesIterator();
        XWPFTable table;
        List<XWPFTableRow> rows;
        List<XWPFTableCell> cells;
        List<XWPFParagraph> paras;
        while (iterator.hasNext()) {
            table = iterator.next();
            rows = table.getRows();
            for (XWPFTableRow row : rows) {
                cells = row.getTableCells();
                for (XWPFTableCell cell : cells) {
                    //虽然paras.size() == 1,但是需要更换单元格指定${}的内容,需要使用XWPFRun对象进行操作、
                    paras = cell.getParagraphs();
                    for (XWPFParagraph para : paras) {
                        this.replaceInPara(para, params);
                    }
                }
            }
        }
    }

    /**
     * 替换docx文件中段落的${}变量
     *
     * @param docx   要替换的docx文件
     * @param params 参数
     */
    public void resetParagraphDOCX(XWPFDocument docx, Map<String, Object> params) {
        for (XWPFParagraph para : docx.getParagraphs()) {
            this.replaceInPara(para, params);
        }
    }

    /**
     * 获取doc、docx文件中表的数据、
     *
     * @param absPath doc、docx文件路径
     * @return List<Map>对象,Map的Key使用 row + "-"+ column 组成、value是单元格内容、
     * @throws IOException IO异常
     */
    public List<LinkedHashMap<String, Object>> getTablesDataList(String absPath) throws IOException {
        if (StringUtils.isEmpty(absPath))
            throw new RuntimeException("文件的绝对路径不能为空");
        if (!absPath.endsWith(".doc") && !absPath.endsWith(".docx") && !absPath.endsWith(".DOC") && !absPath.endsWith(".DOCX"))
            throw new RuntimeException("文件类型必须为'doc'或者'docx'格式");

        if (absPath.endsWith(".doc") || absPath.endsWith(".DOC")) {
            return getDOCTablesDataList(absPath);
        }
        return getDOCXTablesDataList(absPath);

    }


    /**
     * 替换docx文件段落里面的${}变量
     *
     * @param paragraph 要替换的段落
     * @param params    参数
     */
    private void replaceInPara(XWPFParagraph paragraph, Map<String, Object> params) {
        List<XWPFRun> runs;
        Matcher matcher;
        //如果para拆分为XWPFRun的不符合${...}占位符格式,则修改成正确的
        List<XWPFRun> xwpfRuns = this.replaceText(paragraph);
        if (this.matcher(paragraph.getParagraphText()).find()) {
            runs = paragraph.getRuns();//此时得到的runs对象内容应该与xwpfRuns是相同的。
            for (int i = 0; i < runs.size(); i++) {
                XWPFRun run = runs.get(i);
                String runText = run.toString();
                matcher = this.matcher(runText);
                /*
                 *   注意：matcher.find()方法、类似于Iterator.hasNext(),执行一次,获取一次,一次最多只返回一个结果.
                 * */
                if (matcher.find()) {
                    //TODO: 注意*** 此处(matcher = this.matcher(runText))使用runText重新获取matcher对象、否则会出现死循环
                    while ((matcher = this.matcher(runText)).find()) {
                        /*
                         *   matcher.group():  返回 '${name}'
                         *   matcher.group(1): 返回 'name'
                         * */
                        runText = matcher.replaceFirst(String.valueOf(params.get(matcher.group(1))));
                    }
                    //直接调用XWPFRun的setText()方法设置文本时,在底层会重新创建一个XWPFRun,把文本附加在当前文本后面,
                    //所以我们不能直接设值,需要先删除当前run,然后再自己手动插入一个新的run。
                    paragraph.removeRun(i);
                    paragraph.insertNewRun(i).setText(runText);
                }
            }
        }
    }

    /**
     * 合并docx文件当前段落中的runs内容、
     *
     * @param para docx文档的段落
     * @return 当前段落中的文字runs集合、(暂时没有用到返回值)
     */
    private List<XWPFRun> replaceText(XWPFParagraph para) {
        List<XWPFRun> runs = para.getRuns();
        StringBuilder builder = new StringBuilder();
        boolean flag = false;
        for (int i = 0; i < runs.size(); i++) {
            XWPFRun run = runs.get(i);
            String runText = run.toString();
            if (flag || (runText.contains("${") && !runText.contains("}"))) {
                builder.append(runText);
                flag = true;
                para.removeRun(i);
                if (runText.contains("}") && !runText.contains("${")) {
                    flag = false;
                    para.insertNewRun(i).setText(builder.toString());
                    builder = new StringBuilder();
                }
                i--;
            }
        }
        return runs;
    }

    /**
     * 设置docx文件中单元格水平位置和垂直位置
     *
     * @param xwpfTable          Table对象
     * @param verticalLocation   单元格中内容垂直上TOP，下BOTTOM，居中CENTER，BOTH两端对齐
     * @param horizontalLocation 单元格中内容水平居中center,left居左，right居右，both两端对齐
     */
    private void setCellLocation(XWPFTable xwpfTable, String verticalLocation, String horizontalLocation) {
        List<XWPFTableRow> rows = xwpfTable.getRows();
        for (XWPFTableRow row : rows) {
            List<XWPFTableCell> cells = row.getTableCells();
            for (XWPFTableCell cell : cells) {
                CTTc cttc = cell.getCTTc();
                CTP ctp = cttc.getPList().get(0);
                CTPPr ctppr = ctp.getPPr();
                if (ctppr == null) {
                    ctppr = ctp.addNewPPr();
                }
                CTJc ctjc = ctppr.getJc();
                if (ctjc == null) {
                    ctjc = ctppr.addNewJc();
                }
                ctjc.setVal(STJc.Enum.forString(horizontalLocation)); //水平居中
                cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.valueOf(verticalLocation));//垂直居中
            }
        }
    }

    /**
     * 设置docx文件中表格的位置
     *
     * @param xwpfTable Table对象
     * @param location  整个表格居中center,left居左，right居右，both两端对齐
     */
    private void setTableLocation(XWPFTable xwpfTable, String location) {
        CTTbl cttbl = xwpfTable.getCTTbl();
        CTTblPr tblpr = cttbl.getTblPr() == null ? cttbl.addNewTblPr() : cttbl.getTblPr();
        CTJc cTJc = tblpr.addNewJc();
        cTJc.setVal(STJc.Enum.forString(location));
    }


    /**
     * 替换doc文件中所有的${}变量 -->段落和表格
     *
     * @param doc    doc源文件
     * @param params params的key:需要是源文件中${}的变量名、value是填充${}变量的值、
     */
    public void resetAllDOC(HWPFDocument doc, Map<String, Object> params) {
        Range range = doc.getRange();
        params.keySet().forEach(key ->
                range.replaceText("${" + key + "}", params.get(key).toString()));
    }

    /**
     * 替换doc文件中所有表格的${}变量
     *
     * @param doc    doc源文件
     * @param params params的key:需要是源文件中${}的变量名、value是填充${}变量的值、
     */
    public void resetTableDOC(HWPFDocument doc, Map<String, Object> params) {
        Range range = doc.getRange();//此时的Range对象,是文档的页眉和页脚之外的所有内容。
        /*
         *   使用TableIterator tableIterator = new TableIterator(range);来获取Table对象、
         * */
        TableIterator tableIterator = new TableIterator(range);//从当前的Range对象中获取TableIterator、

        while (tableIterator.hasNext()) {
            Table table = tableIterator.next();
            for (int i = 1; i < table.numRows(); i++) {//表头没有变量,不需要遍历、
                TableRow tableRow = table.getRow(i);
                for (String key : params.keySet()) {
                    String placeHolder = "${" + key + "}";
                    tableRow.replaceText(placeHolder, params.get(key).toString());
                }
            }
        }
    }

    /**
     * 替换doc文件中指定表格的${}变量
     *
     * @param doc      doc源文件
     * @param params   params的key:需要是源文件中${}的变量名、value是填充${}变量的值、
     * @param indexArr doc源文件中的表格的index值、
     */
    public void resetTableDOC(HWPFDocument doc, Map<String, Object> params, Integer... indexArr) {
        for (Integer integer : indexArr) {
            if (integer < 0) {
                throw new RuntimeException("index的值不能为负数");
            }
        }
        Range range = doc.getRange();//此时的Range对象,是文档的页眉和页脚之外的所有内容。
        TableIterator tableIterator = new TableIterator(range);//从当前的Range对象中获取TableIterator、

        int tableIndex = 0;
        /*
         *   通过Arrays.asList获取的List、不可以执行list.remove()和list.add()方法、抛出 java.lang.UnsupportedOperationException异常。
         * */
        List<Integer> list = Arrays.asList(indexArr);
        list = new ArrayList<>(list);//单独创建一个新的List对象、
        Collections.sort(list);
        while (tableIterator.hasNext()) {
            if (list.size() == 0) {
                break;
            }
            Table table = tableIterator.next();
            if (tableIndex != list.get(0)) {
                tableIndex++;
                continue;
            }
            for (int i = 0; i < table.numRows(); i++) {//也许表头存在${}变量、少数情况、
                TableRow tableRow = table.getRow(i);
                params.keySet().forEach(key -> tableRow.replaceText("${" + key + "}", params.get(key).toString()));
            }
            list.remove(0);
            tableIndex++;
        }
    }

    /**
     * 正则匹配字符串
     *
     * @param str 需要被匹配的字符串
     * @return Matcher
     */
    private Matcher matcher(String str) {
        // 正则表达式: \$\{(.+?)} --> 匹配：${..} 形式的字符串、不能匹配 ${}。
        Pattern pattern = Pattern.compile("\\$\\{(.+?)}", Pattern.CASE_INSENSITIVE);//忽略大小写
        return pattern.matcher(str);
    }

    /**
     * 处理docx文件表格数据
     *
     * @param absPath docx源文件路径
     * @return data
     * @throws IOException IO
     */
    private List<LinkedHashMap<String, Object>> getDOCXTablesDataList(String absPath) throws IOException {
        List<LinkedHashMap<String, Object>> tableList = new ArrayList<>();
        File file = new File(absPath);
        FileInputStream inputStream = new FileInputStream(file);
        XWPFDocument docx = new XWPFDocument(inputStream);
        List<XWPFTable> tables = docx.getTables();
        for (XWPFTable table : tables) {
            LinkedHashMap<String, Object> tableMap = new LinkedHashMap<>();
            List<XWPFTableRow> tableRows = table.getRows();
            for (int row = 0; row < tableRows.size(); row++) {
                List<XWPFTableCell> tableCells = tableRows.get(row).getTableCells();
                for (int column = 0; column < tableCells.size(); column++) {
                    XWPFTableCell cell = tableCells.get(column);
                    //使用: row-column 作为 key、
                    tableMap.put(row + "-" + column, cell.getText().trim());
                }
            }
            tableList.add(tableMap);
        }
        closeStream(docx, inputStream);
        return tableList;
    }

    /**
     * 处理doc文件表格数据
     *
     * @param absPath doc源文件路径
     * @return data
     * @throws IOException IO
     */
    private List<LinkedHashMap<String, Object>> getDOCTablesDataList(String absPath) throws IOException {
        File file = new File(absPath);
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
                    // tableCell.numParagraphs() == 1; 因此不需要再进行遍历、
                    /*
                     *   HWPFDocument使用(Range、Paragraph、CharacterRun中的)insertBefore()和insertAfter()方法可以修改相应文本信息。
                     *   CharacterRun 类似XWPFRun对象、
                     * */
//                    for (int para = 0; para < tableCell.numParagraphs(); para++) {
//                        Paragraph paragraph = tableCell.getParagraph(para);
//                    }
                }
            }
            tableList.add(tableMap);
        }

        closeStream(document, inputStream);
        return tableList;
    }

    /**
     * 为表格创建空行,并为该行复制同样的表格行属性、
     *
     * @param sourceRow 来源Row
     * @param targetRow 目标Row
     */
    private void copyPropertiesAndEmptyRows(XWPFTableRow sourceRow, XWPFTableRow targetRow) {
        //复制行属性
        targetRow.getCtRow().setTrPr(sourceRow.getCtRow().getTrPr());
        List<XWPFTableCell> cellList = sourceRow.getTableCells();
        if (null == cellList) {
            return;
        }
        //添加列、复制列以及列中段落属性
        XWPFTableCell targetCell;
        for (XWPFTableCell sourceCell : cellList) {
            targetCell = targetRow.addNewTableCell();
            //列属性
            targetCell.getCTTc().setTcPr(sourceCell.getCTTc().getTcPr());
            //段落属性
            targetCell.getParagraphs().get(0).getCTP().setPPr(sourceCell.getParagraphs().get(0).getCTP().getPPr());
        }
    }

    /**
     * 为表格创建非空行,并为该行复制同样的表格行属性。同时完成内容填充
     *
     * @param sourceRow   来源Row
     * @param targetRow   目标Row
     * @param tableHeader 需要被填充的列,Tips:** tableHeader 需要 与 params的key相同。
     * @param param       填充的数据
     */
    private void copyPropertiesAndNonEmptyRows(XWPFTableRow sourceRow, XWPFTableRow targetRow, List<String> tableHeader, Map<String, Object> param) {
        //复制行属性
        targetRow.getCtRow().setTrPr(sourceRow.getCtRow().getTrPr());
        List<XWPFTableCell> cellList = sourceRow.getTableCells();
        if (cellList == null) {
            return;
        }
        //添加列、复制列以及列中段落属性
        XWPFTableCell targetCell;
        for (int i = 0; i < cellList.size(); i++) {
            targetCell = targetRow.addNewTableCell();
            //列属性
            targetCell.getCTTc().setTcPr(cellList.get(i).getCTTc().getTcPr());

            //段落属性
            List<XWPFParagraph> paragraphList = cellList.get(i).getParagraphs();
            if (isNotEmpty(paragraphList)) {
                targetCell.getParagraphs().get(0).getCTP().setPPr(paragraphList.get(0).getCTP().getPPr());
                List<XWPFRun> xwpfRunList = paragraphList.get(0).getRuns();
                if (isNotEmpty(xwpfRunList)) {
                    //字体属性
                    XWPFRun cellR = targetCell.getParagraphs().get(0).createRun();
                    cellR.setBold(xwpfRunList.get(0).isBold());
                    cellR.setFontFamily(xwpfRunList.get(0).getFontFamily());
                    cellR.setFontSize(xwpfRunList.get(0).getFontSize());
                    //TODO: 设置单元格内容(需要在此处设置内容,才可以copy字体样式)
                    cellR.setText(param.get(tableHeader.get(i)).toString());
                } else
                    targetCell.setText(param.get(tableHeader.get(i)).toString());
            } else
                targetCell.setText(param.get(tableHeader.get(i)).toString());
        }
    }

    /**
     * @param table   docx文件中的表格
     * @param add     增加或删除行数 if add>0 增加行 add<0 删除行
     * @param fromRow 添加开始行位置(fromRow-1是模版行),from >= 1、不允许复制第一行
     */
    private void insertOrRemoveEmptyRows(XWPFTable table, int add, int fromRow) {
        int size = table.getRows().size();
        if (add == 0 || fromRow > size || size < 1 || fromRow < 1)
            throw new RuntimeException("add参数不可等于0,fromRow参数必须大于0或fromRow超出表格总数");
        XWPFTableRow row = table.getRow(fromRow - 1);
        if (add > 0) {
            while (add > 0) {
                /*
                 *   执行换为多线程copy、
                 * */
                Executors.newCachedThreadPool().execute(() -> copyPropertiesAndEmptyRows(row, table.insertNewTableRow(fromRow)));
//                copyPropertiesAndEmptyRows(row, table.insertNewTableRow(fromRow));
                add--;
            }
        } else {
            while (add < 0) {
                table.removeRow(fromRow - 1);
                add++;
            }
        }
    }

    /**
     * 不限制插入行数和插入位置、
     *
     * @param table       docx文件中的表格
     * @param add         增加或删除行数 if add>0 增加行 add<0 删除行
     * @param fromRow     添加开始行位置(fromRow-1是模版行),from >= 1、不允许复制第一行
     * @param tableHeader 表头的名称(第一行的单元格内容)
     *                    Tips:** tableHeader的值 需要 与 params的key相同。
     * @param params      新增行的内容
     */
    private void insertNotEmptyRows(XWPFTable table, int add, int fromRow, List<String> tableHeader, List<Map<String, Object>> params) {
        int size = table.getRows().size();
        if (add <= 0 || fromRow > size || size <= 1 || fromRow < 1) //不允许复制表头属性
            throw new RuntimeException("add参数不可等于0,fromRow参数必须大于0或fromRow超出表格总数");
        XWPFTableRow row = table.getRow(fromRow - 1);
        int index = params.size() - 1;
        //换为多线程的方式来执行copy新增、
        while (add > 0) {
            if (add > params.size()) {
                Executors.newCachedThreadPool().execute(() -> copyPropertiesAndEmptyRows(row, table.insertNewTableRow(fromRow)));
//                copyPropertiesAndEmptyRows(row, table.insertNewTableRow(fromRow));
                add--;
                continue;
            }
            Map<String, Object> param = params.get(index);
            Executors.newCachedThreadPool().execute(() -> copyPropertiesAndNonEmptyRows(row, table.insertNewTableRow(fromRow), tableHeader, param));
//            copyPropertiesAndNonEmptyRows(row, table.insertNewTableRow(fromRow), tableHeader, param);

            index--;
            add--;
        }
    }

    /**
     * 限制插入行数(paras.size()+1)和插入位置(表的最后一行)
     *
     * @param table       docx文件中的表格
     * @param tableHeader 表头的名称(第一行的单元格内容)
     *                    Tips:** tableHeader 需要 与 params的key相同。
     * @param params      新增行的内容、有多少条数据,就新增多少行
     */
    private void insertNotEmptyRows(XWPFTable table, List<String> tableHeader, List<Map<String, Object>> params) {
        int index = params.size() - 1;//如果使用index=0; index++ -->填充的数据顺序会变为逆序、
        int fromRow = table.getRows().size();
//        int size = table.getRows().size();
        if (fromRow <= 1) //不允许复制表头属性
            throw new RuntimeException("table至少需要有2行");
        XWPFTableRow row = table.getRow(fromRow - 1);
        while (index >= 0) {
            /*
             *   先复制属性、再填充内容,会出现BUG。而且效率更低、
             * */
//            copyPropertiesAndEmptyRows(row, table.insertNewTableRow(fromRow));
//            //得到新增的空行、
//            XWPFTableRow newRow = table.getRow(size++);
//            //填充数据
//            Map<String, Object> param = params.get(count);
//            for (int i = 0; i < newRow.getTableCells().size(); i++) {
//                XWPFTableCell cell = newRow.getCell(i);
//                /*
//                 *   TODO: 两种方式都无法解决 --> 插入n条数据时，前面n-1条的行都是空、第n行显示全部的数据、
//                 *      1、为XWPFTableCell创建XWPFRun对象后,再setText
//                 *      2、直接XWPFTableCell对象setText
//                 * */
//                /*
//                *   1、 for (int j = 0; j < cell.getParagraphs().size(); j++) {
//                        XWPFParagraph xwpfParagraph = cell.getParagraphs().get(j);
//                        XWPFRun xwpfRun = xwpfParagraph.createRun();
//                        xwpfRun.setText(param.get(tableHeader.get(i)).toString());
//                }
//                * */
//                /*
//                *   2、 cell.setText(param.get(tableHeader.get(i)).toString());
//                * */
//
//            }
            /*
             *   换为多线程方式copy属性、
             * */
            int finalIndex = index;
            Executors.newCachedThreadPool().execute(() -> copyPropertiesAndNonEmptyRows(row, table.insertNewTableRow(fromRow), tableHeader, params.get(finalIndex)));
            //copyPropertiesAndNonEmptyRows(row, table.insertNewTableRow(fromRow), tableHeader, params.get(index));

            index--;
        }
    }

    public void closeStream(HWPFDocument document, OutputStream outputStream, InputStream inputStream) throws IOException {
        if (outputStream != null)
            outputStream.close();
        if (document != null) {
            document.close();
        }
        if (inputStream != null)
            inputStream.close();
    }

    public void closeStream(HWPFDocument document, InputStream inputStream) throws IOException {
        if (document != null)
            document.close();
        if (inputStream != null)
            inputStream.close();
    }

    public void closeStream(HWPFDocument document, OutputStream outputStream) throws IOException {
        if (document != null)
            document.close();
        if (outputStream != null)
            outputStream.close();
    }

    public void closeStream(XWPFDocument document, OutputStream outputStream, InputStream inputStream) throws IOException {
        if (outputStream != null)
            outputStream.close();
        if (document != null) {
            document.close();
        }
        if (inputStream != null)
            inputStream.close();
    }

    public void closeStream(XWPFDocument document, InputStream inputStream) throws IOException {
        if (document != null)
            document.close();
        if (inputStream != null)
            inputStream.close();
    }

    public void closeStream(XWPFDocument document, OutputStream outputStream) throws IOException {
        if (document != null)
            document.close();
        if (outputStream != null)
            outputStream.close();
    }

    private boolean isNotEmpty(Collection collection) {
        return collection != null && collection.size() > 0;
    }
}
