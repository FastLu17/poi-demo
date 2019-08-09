package com.lxf.poi.util;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.*;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * @author 小66
 * @Description
 * @create 2019-08-09 15:17
 **/
public class POIUtils {

    private final String BASE_URL = "C:\\Users\\Administrator\\Desktop\\POI\\";

    public String createNewRowsForTable(String filePath, List<String> tableHeader, List<Map<String, Object>> params) throws IOException {
        createNewRowsForTable(filePath, tableHeader, params, 0);
        return "";
    }

    public String createNewRowsForTable(String filePath, List<String> tableHeader, List<Map<String, Object>> params, int... tableIndexArr) throws IOException {
        if (!filePath.endsWith(".docx") || !filePath.endsWith(".DOCX"))
            throw new RuntimeException("无法处理非docx文件");
        FileInputStream inputStream = new FileInputStream(filePath);
        XWPFDocument docx = new XWPFDocument(inputStream);
        List<XWPFTable> tables = docx.getTables();
        for (int tableIndex : tableIndexArr) {
            if (tableIndex >= tables.size()) {
                throw new RuntimeException("文件中表格数量低于" + tableIndex + 1 + "个");
            }
            XWPFTable table = tables.get(tableIndex);
            //新增rows、
        }

        return "";
    }

    /**
     * @param fileName    需要创建的文件名
     * @param tableName   文件中的表名
     * @param tableHeader 表格的表头
     * @param params      表格的内容
     * @return 文件路径
     * @throws IOException IO
     */
    public String dataInsertIntoTable(String fileName, String tableName, List<String> tableHeader, List<Map<String, Object>> params) throws IOException {

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
        width.setW(BigInteger.valueOf(8000));
        //设置表格宽带为固定、
        width.setType(STTblWidth.DXA);//STTblWidth.AUTO:自动-->宽度无效.

        //TODO: 如何设置表格位置?

        List<XWPFTableRow> tableRows = table.getRows();

        //获取表头
        XWPFTableRow header = tableRows.get(0);
        List<XWPFTableCell> tableCells = header.getTableCells();
        //设置表头样式和属性
        for (int i = 0; i < tableHeader.size(); i++) {
            XWPFTableCell headerCell = tableCells.get(i);
            //垂直居中
            headerCell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
            XWPFParagraph headerPara = headerCell.addParagraph();//获取单元格的段落、
            XWPFRun headerRun = headerPara.createRun();//XWPFRun是最小的单位.(文本)
            headerRun.setText(tableHeader.get(i));
            headerRun.setBold(true);
            headerRun.setFontSize(16);
            headerRun.setFontFamily("SimHei");//黑体
            headerRun.setColor("66FF66");
            headerRun.setShadow(true);//阴影
            //水平居中
            headerPara.setAlignment(ParagraphAlignment.CENTER);
            //垂直居中
            headerPara.setVerticalAlignment(TextAlignment.CENTER);
        }

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
                cellRun.setFontSize(12);
                cellRun.setFontFamily("宋体");//黑体
                cellRun.setUnderline(UnderlinePatterns.SINGLE);

                cellParagraph.setVerticalAlignment(TextAlignment.CENTER);
                cellParagraph.setAlignment(ParagraphAlignment.CENTER);
            }
        }

        FileOutputStream outputStream = new FileOutputStream(filePath);
        docx.write(outputStream);
        outputStream.close();

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
     * 获取doc文件中表的数据、
     *
     * @param absPath doc文件路径
     * @return List<Map>对象,Map的Key使用 row + "-"+ column 组成、value是单元格内容、
     * @throws IOException IO异常
     */
    public List<LinkedHashMap<String, Object>> getDocTablesList(String absPath) throws IOException {
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
     * 替换docx文件段落里面的变量
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
}
