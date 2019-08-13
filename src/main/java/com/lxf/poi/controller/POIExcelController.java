package com.lxf.poi.controller;

import com.lxf.poi.util.POIExcelUtil;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.hssf.extractor.ExcelExtractor;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.ss.util.RegionUtil;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.net.URL;
import java.util.HashMap;
import java.util.Map;

/**
 * @author 小66
 * @create 2019-08-12 21:43
 **/
@RestController
@Slf4j
public class POIExcelController {

    @Autowired
    private POIExcelUtil excelUtil;

    private final String BASE_DIRECTORY_PATH = "C:\\Users\\Administrator\\Desktop\\POI\\";
    private final String XLS_TEMPLATE_FILE_PATH = BASE_DIRECTORY_PATH + "HSSF测试模板.xls";
    private final String XLSX_TEMPLATE_FILE_PATH = BASE_DIRECTORY_PATH + "XSSF测试模板.xlsx";

    @GetMapping("createXls")
    public String createXls() throws Exception {
//        POIFSFileSystem system = new POIFSFileSystem(new File(BASE_DIRECTORY_PATH + "HSSF测试read.xls"));
//        Workbook book = WorkbookFactory.create(system);
//        HSSFWorkbook workbook = new HSSFWorkbook(system);  //多种方式可以创建Workbook对象、
        HSSFWorkbook workbook = new HSSFWorkbook();

        HSSFDataFormat format = workbook.createDataFormat();//格式化

        HSSFSheet sheet = workbook.createSheet("sheet");
        sheet.setSelected(false);
        HSSFSheet sheetSelected = workbook.createSheet("sheetSelected");
        sheetSelected.setSelected(true);//设置默认被选中、没生效

        CellRangeAddress range = CellRangeAddress.valueOf("C5:F20");//不同的方式创建range、
        CellRangeAddress rangeAddress = new CellRangeAddress(1, 1, 1, 1);
        sheet.setAutoFilter(rangeAddress);//设置自动过滤、

        //设置缩放比例
        sheet.setZoom(115);// 表示:115%

        sheet.createFreezePane(1, 1);//设置冻结窗格--行和列(不受滚动影响)

        //设置头尾(没效果)
        HSSFHeader header = sheet.getHeader();
        HSSFFooter footer = sheet.getFooter();
        header.setRight(HSSFHeader.font("Stencil-Normal", "Italic") +
                HSSFHeader.fontSize((short) 16) + "Right w/ Stencil-Normal Italic font and size 16");
        footer.setRight("Footer Right");

        Row row = sheet.createRow(1);
        /*
         *   RowHeightInPoints = 12.75
         *   RowHeight = 255
         *   20倍的关系、
         * */
        float defaultHeightInPoints = sheet.getDefaultRowHeightInPoints();//获取默认行高、
        sheet.autoSizeColumn(1);//设置某一列自根据内容自动调整宽度、
        row.setHeightInPoints(2 * defaultHeightInPoints);//设置两倍行高、

        //设置打印区域
        workbook.setPrintArea(0, 0, 9, 0, 9);

        Font font = getFont(workbook, 200, "Consolas", IndexedColors.RED.index);
        Font font2 = getFont(workbook, 300, "黑体", IndexedColors.GREEN.index);

        //富文本、
        HSSFRichTextString richString = new HSSFRichTextString("Hello, World!");
        richString.applyFont(0, 6, font);
        richString.applyFont(6, 13, font2);

        createCell(workbook, font, row, 0, HorizontalAlignment.CENTER, VerticalAlignment.BOTTOM, "Align It");
        createCell(workbook, font, row, 1, HorizontalAlignment.CENTER_SELECTION, VerticalAlignment.BOTTOM, "Align It Align It");
        createCell(workbook, font, row, 2, HorizontalAlignment.FILL, VerticalAlignment.CENTER, "Align It");
        createCell(workbook, font, row, 3, HorizontalAlignment.GENERAL, VerticalAlignment.CENTER, "Align It");
        createCell(workbook, font, row, 4, HorizontalAlignment.JUSTIFY, VerticalAlignment.JUSTIFY, "Align It");
        createCell(workbook, font, row, 5, HorizontalAlignment.LEFT, VerticalAlignment.TOP, "Align It");
        createCell(workbook, font, row, 6, HorizontalAlignment.RIGHT, VerticalAlignment.TOP, "Align It");
        Cell cell = row.createCell(7);
        cell.setCellValue(richString);

        row.setZeroHeight(false);//是否设置此行高度为0、隐藏(hidden)


        FileOutputStream outputStream = new FileOutputStream(BASE_DIRECTORY_PATH + "HSSF测试create.xls");
        workbook.write(outputStream);

        excelUtil.closeStream(workbook, outputStream);

        return BASE_DIRECTORY_PATH + "HSSF测试create.xls";
    }

    @GetMapping("readXls")
    public String readXls() throws Exception {
        POIFSFileSystem system = new POIFSFileSystem(new File(XLS_TEMPLATE_FILE_PATH));
        Workbook book = WorkbookFactory.create(system);
//        simpleIterator(book);
        simpleIterator(book);
        book.close();
        return "...";
    }

    @GetMapping("getXlsText")
    public String getXlsText() throws Exception {
        POIFSFileSystem system = new POIFSFileSystem(new File(XLS_TEMPLATE_FILE_PATH));
        HSSFWorkbook workbook = new HSSFWorkbook(system);
        //Text Extraction：文字抽取
        ExcelExtractor extractor = new ExcelExtractor(workbook);

        extractor.setFormulasNotResults(false);//设置公式有无返回值--> true：输出公式、 false：输出结果
        extractor.setIncludeSheetNames(false);
        return extractor.getText();
    }

    @GetMapping("mergingCells")
    public void createMergingCells() throws Exception {
        mergingCells(new CellRangeAddress(1, 1, 0, 1), XLS_TEMPLATE_FILE_PATH);

    }

    /**
     * 条件格式化、ConditionalFormatting
     *
     * @param sheet
     */
    public void formating(Sheet sheet) {
        SheetConditionalFormatting sheetCF = sheet.getSheetConditionalFormatting();

        ConditionalFormattingRule rule1 = sheetCF.createConditionalFormattingRule(ComparisonOperator.EQUAL, "0");
        ConditionalFormattingRule rule2 = sheetCF.createConditionalFormattingRule(ComparisonOperator.BETWEEN, "0", "10");
        FontFormatting fontFmt = rule1.createFontFormatting();
        fontFmt.setFontStyle(true, false);
        fontFmt.setFontColorIndex(IndexedColors.DARK_RED.index);
    }

    /**
     * 展示HSSF工具类的使用、
     * HSSFRegionUtil和HSSFCellUtil被弃用、
     */
    @GetMapping("utils")
    public void utils() throws Exception {
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet();

        CellRangeAddress rangeAddress = CellRangeAddress.valueOf("C5:F20");
//        CellRangeAddress rangeAddress = new CellRangeAddress(1, 9, 1, 9);
        RegionUtil.setBorderBottom(BorderStyle.THIN.getCode(), rangeAddress, sheet);
        RegionUtil.setBottomBorderColor(IndexedColors.RED.index, rangeAddress, sheet);
        RegionUtil.setBorderRight(BorderStyle.MEDIUM.getCode(), rangeAddress, sheet);
        RegionUtil.setRightBorderColor(IndexedColors.GREEN.index, rangeAddress, sheet);
        RegionUtil.setBorderTop(BorderStyle.THIN.getCode(), rangeAddress, sheet);
        RegionUtil.setTopBorderColor(IndexedColors.RED.index, rangeAddress, sheet);
        RegionUtil.setBorderLeft(BorderStyle.MEDIUM.getCode(), rangeAddress, sheet);
        RegionUtil.setLeftBorderColor(IndexedColors.GREEN.index, rangeAddress, sheet);

        CellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);

        HSSFRow row = sheet.createRow(0);
        HSSFRow row1 = sheet.createRow(1);
        Cell cell_00 = CellUtil.createCell(row, 0, "CellUtil--00", style);
        Cell cell_01 = CellUtil.createCell(row, 1, "CellUtil--01", style);
        Cell cell_10 = CellUtil.createCell(row1, 0, "CellUtil--10", style);
        Cell cell_11 = CellUtil.createCell(row1, 1, "CellUtil--11", style);

        FileOutputStream outputStream = new FileOutputStream(BASE_DIRECTORY_PATH + "HSSF测试utils.xls");
        workbook.write(outputStream);

        excelUtil.closeStream(workbook, outputStream);
    }

    /**
     * 画简单的形状、常用来插入图片、
     *
     * @throws Exception
     */
    @GetMapping("drawShapes")
    public void drawShapes() throws Exception {
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet();
        //画图的顶级管理器(元老)，一个sheet只能获取一个、
        HSSFPatriarch patriarch = sheet.createDrawingPatriarch();

        //创建锚对象： 255:单元格默认高度、 1023:单元格默认宽带、
        HSSFClientAnchor anchorLine = new HSSFClientAnchor(0, 0, 1023, 255, (short) 1, 0, (short) 2, 1);
        HSSFSimpleShape shape = patriarch.createSimpleShape(anchorLine);
        shape.setShapeType(HSSFSimpleShape.OBJECT_TYPE_LINE);

        /*
         *   画椭圆
         * */
        HSSFClientAnchor anchorOval = new HSSFClientAnchor(0, 0, 1023, 255, (short) 1, 3, (short) 2, 4);
        HSSFSimpleShape shapeOval = patriarch.createSimpleShape(anchorOval);
        shapeOval.setShapeType(HSSFSimpleShape.OBJECT_TYPE_OVAL);

        /*
         *   输入框
         * */
        HSSFClientAnchor anchorTextBox = new HSSFClientAnchor(0, 0, 1023, 255, (short) 4, 1, (short) 5, 2);
        HSSFTextbox textBox = patriarch.createTextbox(anchorTextBox);
        textBox.setLineStyle(HSSFShape.LINESTYLE_SOLID);
        HSSFRichTextString richTextString = new HSSFRichTextString("This is a test");
        richTextString.applyFont(getFont(workbook, 200, "Consolas", IndexedColors.RED.index));
        textBox.setString(richTextString);
        textBox.setVerticalAlignment(HSSFTextbox.VERTICAL_ALIGNMENT_CENTER);
        textBox.setHorizontalAlignment(HSSFTextbox.VERTICAL_ALIGNMENT_CENTER);
        textBox.setFillColor(IndexedColors.GREEN.index);//背景色始终为黑色,不设置就是白色、

        //插入图片
        URL url = new URL("http://p20.qhimgs3.com/dr/240_240_/t011c15f8c12e731c01.jpg?t=1558300397");
        BufferedImage read = ImageIO.read(url);
        ByteArrayOutputStream byteArrayOS = new ByteArrayOutputStream();
        ImageIO.write(read, "JPEG", byteArrayOS);
        byte[] bytes = byteArrayOS.toByteArray();

        int pictureIndex = workbook.addPicture(bytes, Workbook.PICTURE_TYPE_JPEG);
        HSSFClientAnchor anchorPicture = new HSSFClientAnchor(0, 0, 1023, 255, (short) 7, 1, (short) 9, 2);
        patriarch.createPicture(anchorPicture, pictureIndex);
        read.flush();

        FileOutputStream outputStream = new FileOutputStream(BASE_DIRECTORY_PATH + "HSSF测试drawShapes.xls");
        workbook.write(outputStream);

        excelUtil.closeStream(workbook, outputStream);
    }

    /**
     * 插入图片到Excel中、
     *
     * @throws Exception
     */
    @GetMapping("pictures")
    public void pictures() throws Exception {
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet();
        //画图的顶级管理器(元老)，一个sheet只能获取一个、
        HSSFPatriarch patriarch = sheet.createDrawingPatriarch();
        //插入图片
        URL url = new URL("http://p20.qhimgs3.com/dr/240_240_/t011c15f8c12e731c01.jpg?t=1558300397");
        BufferedImage read = ImageIO.read(url);
        ByteArrayOutputStream byteArrayOS = new ByteArrayOutputStream();
        ImageIO.write(read, "JPEG", byteArrayOS);
        byte[] bytes = byteArrayOS.toByteArray();

        int pictureIndex = workbook.addPicture(bytes, Workbook.PICTURE_TYPE_JPEG);
        HSSFClientAnchor anchorPicture = new HSSFClientAnchor(0, 0, 1023, 255, (short) 7, 1, (short) 8, 3);
        //开始画图片、
        patriarch.createPicture(anchorPicture, pictureIndex);
        read.flush();

        FileOutputStream outputStream = new FileOutputStream(BASE_DIRECTORY_PATH + "HSSF测试pictures.xls");
        workbook.write(outputStream);

        excelUtil.closeStream(workbook, outputStream);
    }

    /**
     *  超链接引入图片、
     */
    @GetMapping("hyperlink")
    public void hyperlink()throws Exception{
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFHyperlink hyperlink = workbook.getCreationHelper().createHyperlink(HyperlinkType.URL);
        hyperlink.setShortFilename("驴子");
        hyperlink.setAddress("http://p20.qhimgs3.com/dr/240_240_/t011c15f8c12e731c01.jpg?t=1558300397");
        HSSFSheet sheet = workbook.createSheet();
        HSSFCell cell = sheet.createRow(0).createCell(0);
        HSSFCellStyle style = workbook.createCellStyle();
        Font font = getFont(workbook, 200, "Consolas", IndexedColors.BLUE.index);
        style.setFont(font);
        cell.setHyperlink(hyperlink);
        cell.setCellValue("图片链接");

        FileOutputStream outputStream = new FileOutputStream(BASE_DIRECTORY_PATH + "HSSF测试hyperlink.xls");
        workbook.write(outputStream);

        excelUtil.closeStream(workbook, outputStream);
    }

    @GetMapping("createFreezePane")
    public void createFreezePane() throws Exception {
        Workbook wb = new HSSFWorkbook();
        Sheet sheet1 = wb.createSheet();
        Sheet sheet2 = wb.createSheet();
        Sheet sheet3 = wb.createSheet();

        //设置固定行和列的窗格、(滚动一直会展示,多用于表头)
        sheet1.createFreezePane(0, 1, 0, 1);//固定行、
        sheet2.createFreezePane(1, 0, 1, 0);//固定列、
        sheet3.createFreezePane(2, 2);

        FileOutputStream outputStream = new FileOutputStream(BASE_DIRECTORY_PATH + "HSSF测试freezePane.xls");
        wb.write(outputStream);

        excelUtil.closeStream(wb, outputStream);
    }

    @GetMapping("setCellProperties")
    public void setCellProperties() throws Exception {
        Workbook workbook = new HSSFWorkbook();
        Sheet sheet = workbook.createSheet();
        Map<String, Object> properties = new HashMap<>();

        properties.put(CellUtil.BORDER_TOP, BorderStyle.MEDIUM);
        properties.put(CellUtil.BORDER_BOTTOM, BorderStyle.MEDIUM);
        properties.put(CellUtil.BORDER_LEFT, BorderStyle.MEDIUM);
        properties.put(CellUtil.BORDER_RIGHT, BorderStyle.MEDIUM);

        properties.put(CellUtil.TOP_BORDER_COLOR, IndexedColors.RED.index);
        properties.put(CellUtil.BOTTOM_BORDER_COLOR, IndexedColors.RED.index);
        properties.put(CellUtil.LEFT_BORDER_COLOR, IndexedColors.RED.index);
        properties.put(CellUtil.RIGHT_BORDER_COLOR, IndexedColors.RED.index);

        // 应用Border属性到Cell、
        Row row = sheet.createRow(1);
        Cell cell1 = CellUtil.createCell(row, 1, "Hello");
        Cell cell2 = CellUtil.createCell(row, 2, "World");
        /*
         *   通过Properties设置的属性、会与现有的单元格属性合并在一起,
         *       如果有相同的属性,则替换为Properties中的属性、
         * */
        CellUtil.setCellStyleProperties(cell1, properties);
        CellUtil.setCellStyleProperties(cell2, properties);

        FileOutputStream outputStream = new FileOutputStream(BASE_DIRECTORY_PATH + "HSSF测试cellStyleProperties.xls");
        workbook.write(outputStream);

        excelUtil.closeStream(workbook, outputStream);
    }

    private void createCell(Workbook wb, Font font, Row row, int column, HorizontalAlignment halign, VerticalAlignment valign, String value) {
        Cell cell = row.createCell(column);
        cell.setCellValue(value);

        //指定样式、
        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setAlignment(halign);
        cellStyle.setVerticalAlignment(valign);

        //设置边框属性
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBottomBorderColor(IndexedColors.YELLOW.getIndex());
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setLeftBorderColor(IndexedColors.GREEN.getIndex());
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setRightBorderColor(IndexedColors.BLUE.getIndex());
        cellStyle.setBorderTop(BorderStyle.MEDIUM_DASHED);
        cellStyle.setTopBorderColor(IndexedColors.RED.getIndex());

        //设置填充属性
//        cellStyle.setFillBackgroundColor(IndexedColors.AQUA.getIndex());
//        cellStyle.setFillPattern(FillPatternType.BIG_SPOTS);

        //设置字体属性
        cellStyle.setFont(font);

        //设置换行
        cellStyle.setWrapText(true);

        cell.setCellStyle(cellStyle);
    }

    /**
     * 遍历Excel表格、
     * row.getCell(0, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);//指定当前单元格为null and blank cells时的策略、(暂时不知道有何用)
     *
     * @param workbook
     */
    public void simpleIterator(Workbook workbook) {
        DataFormatter formatter = new DataFormatter();
        for (Sheet sheet : workbook) {//sheetIterator
            int lastRowNum = sheet.getLastRowNum();
            log.info("lastRowNum = {}", lastRowNum);//获取最后一行的位置、
            for (Row row : sheet) {//rowIterator
                short lastCellNum = row.getLastCellNum();
                log.info("lastCellNum = {}", lastCellNum);//获取最后一列的位置、
                for (Cell cell : row) {//cellIterator
//                    cell.getStringCellValue();
//                    cell.getNumericCellValue();
                    CellReference cellRef = new CellReference(row.getRowNum(), cell.getColumnIndex());
                    log.info("cellRef.formatAsString = {}", cellRef.formatAsString());
                    String formatValue = formatter.formatCellValue(cell);//不进行formatter、则需要对单元格内容进行判断后再获取、
                    CellAddress address = cell.getAddress();//获取当前单元格的坐标
                    log.info("row = {}, column = {}, value = {}", address.getRow(), address.getColumn(), formatValue);
                }
            }
        }
    }

    public void mergingCells(CellRangeAddress rangeAddress, String path) throws Exception {
        FileInputStream inputStream = new FileInputStream(path);
        HSSFWorkbook workbook = new HSSFWorkbook(inputStream);
        Sheet sheet = workbook.getSheetAt(0);
        sheet.addMergedRegionUnsafe(rangeAddress);//两个单元格都有值,都会保留,但是只显示一个单元格的值、
        FileOutputStream outputStream = new FileOutputStream(path);
        workbook.write(outputStream);
        excelUtil.closeStream(workbook, outputStream, inputStream);
    }

    /**
     * 不要循环创建字体样式,尽量重用、
     *
     * @param workbook
     * @param fontHeight 类似字体大小、但是与fontSize不一样。 200 height <==> 10 size
     * @param fontName   字体名称
     * @param color      eg:IndexedColors.GREEN.index,HSSFColor.RED.index
     * @return
     */
    public Font getFont(Workbook workbook, int fontHeight, String fontName, short color) {
        Font font = workbook.createFont();
//        font.setFontHeightInPoints((short)500);//目前没看到效果
        font.setFontName(fontName);
        font.setBold(true);
        font.setFontHeight((short) fontHeight);
        font.setItalic(true);//斜体
//        font.setStrikeout(true);//删除线
        font.setUnderline(Font.U_SINGLE);

        //设置字体颜色
        font.setColor(color);

        return font;
    }
}
