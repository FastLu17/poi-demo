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
import org.springframework.util.MimeTypeUtils;
import org.springframework.util.concurrent.ListenableFuture;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.*;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.*;

/**
 * @author 小66
 * @create 2019-08-12 21:43
 **/
@RestController
@Slf4j
public class POIExcelController {

    private final POIExcelUtil excelUtil;

    private final String BASE_DIRECTORY_PATH = "C:\\Users\\Administrator\\Desktop\\POI\\";
    private final String XLS_TEMPLATE_FILE_PATH = BASE_DIRECTORY_PATH + "HSSF测试模板.xls";

    @Autowired
    public POIExcelController(POIExcelUtil excelUtil) {
        this.excelUtil = excelUtil;
    }

    @GetMapping("createXls")
    public String createXls() throws Exception {
//        POIFSFileSystem system = new POIFSFileSystem(new File(BASE_DIRECTORY_PATH + "HSSF测试read.xls"));
//        Workbook book = WorkbookFactory.create(system);
//        HSSFWorkbook workbook = new HSSFWorkbook(system);  //多种方式可以创建Workbook对象、
        HSSFWorkbook workbook = new HSSFWorkbook();

        HSSFSheet sheet = workbook.createSheet("sheet");
        sheet.setSelected(true);//设置默认被选中、没生效

        HSSFSheet selectedSheet = workbook.createSheet("selectedSheet");
        selectedSheet.setDisplayGridlines(false);//隐藏Excel网格线,默认值为true
        selectedSheet.setGridsPrinted(true);//打印时显示网格线,默认值为false
        workbook.setActiveSheet(1);//设置默认工作表

        CellRangeAddress range = CellRangeAddress.valueOf("C5:F20");//不同的方式创建range、
        CellRangeAddress rangeAddress = new CellRangeAddress(1, 1, 1, 1);
        selectedSheet.setAutoFilter(rangeAddress);//设置自动过滤、

        //设置缩放比例
        selectedSheet.setZoom(115);// 表示:115%

        selectedSheet.createFreezePane(1, 1);//设置冻结窗格--行和列(不受滚动影响)

        //设置头尾(没效果-->应该是Excel设置了不展示头尾)
        HSSFHeader header = selectedSheet.getHeader();
        HSSFFooter footer = selectedSheet.getFooter();
        header.setRight(HSSFHeader.font("Stencil-Normal", "Italic") +
                HSSFHeader.fontSize((short) 16) + "Right w/ Stencil-Normal Italic font and size 16");
        footer.setRight("Footer Right");

        Row row = selectedSheet.createRow(1);
        /*
         *   RowHeightInPoints = 12.75(磅)
         *   RowHeight = 255
         *   20倍的关系、
         *   defaultColumnWidth = 8(字符)  1字符≈5.55磅、
         * */
        float defaultHeightInPoints = selectedSheet.getDefaultRowHeightInPoints();//获取默认行高、
        selectedSheet.autoSizeColumn(1);//设置某一列自根据内容自动调整宽度、(默认不会自动换行,需要设置style.setWrapText(true);)
        row.setHeightInPoints(2 * defaultHeightInPoints);//设置两倍行高、
        int defaultColumnWidth = selectedSheet.getDefaultColumnWidth();//8 -->8个字符、
        System.out.println("defaultColumnWidth = " + defaultColumnWidth);

        //TODO: 这个参数的单位是1/256个字符宽度 -->表示设置为3个字符宽度、
        selectedSheet.setColumnWidth(1, 3 * 256);

        //设置打印区域
        workbook.setPrintArea(0, 0, 9, 0, 9);

        Font font = excelUtil.getFont(workbook, IndexedColors.RED.index, false, null);
        Font font2 = excelUtil.getFont(workbook, IndexedColors.GREEN.index, false, null);

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
        HSSFWorkbook book = (HSSFWorkbook) WorkbookFactory.create(system);
        HSSFSheet sheet = book.getSheetAt(0);
        Map<String, PictureData> pictureDataMap = getPictureDataMap(sheet);
        List<Map<String, Object>> dataList = new ArrayList<>();
        simpleIterator(book, dataList, pictureDataMap);
        System.out.println("dataList = " + dataList);
        book.close();
        return "...";
    }

    private Map<String, PictureData> getPictureDataMap(HSSFSheet sheet) {
        HSSFPatriarch patriarch = sheet.createDrawingPatriarch();
//        List<HSSFPictureData> hssfPictureDataList = book.getAllPictures(); //无法定位Picture、
        Map<String, PictureData> pictureDataMap = new HashMap<>();
        List<HSSFShape> shapeList = patriarch.getChildren();
        for (HSSFShape hssfShape : shapeList) {
            if (hssfShape instanceof HSSFPicture) {
                HSSFPicture picture = (HSSFPicture) hssfShape;
                HSSFClientAnchor anchor = picture.getClientAnchor();
                short col1 = anchor.getCol1();
                short col2 = anchor.getCol2();
                int row1 = anchor.getRow1();
                int row2 = anchor.getRow2();
                int dx2 = anchor.getDx2();
                int dy2 = anchor.getDy2();
                //只取图片在一个单元格内的、
                if ((col1 == col2 && row1 == row2) || (row1 + 1 == row2 && col1 + 1 == col2 && dx2 == 0 && dy2 == 0)) {
                    pictureDataMap.put(row1 + "-" + col1, picture.getPictureData());
                }
            }
        }
        return pictureDataMap;
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
        FileInputStream inputStream = new FileInputStream(BASE_DIRECTORY_PATH + "HSSF测试method.xls");
        HSSFWorkbook workbook = new HSSFWorkbook(inputStream);
        HSSFSheet sheet = workbook.getSheetAt(0);

        excelUtil.mergingCells(sheet, new CellRangeAddress(1, 1, 0, 1));

        excelUtil.closeStream(workbook, inputStream);

    }

    /**
     * 测试获取CellValue的不同方式：
     * 1、excelUtil.getObjectCellValue(cell);
     * 2、excelUtil.getStringCellValue(cell);
     *
     * @throws Exception
     */
    @GetMapping("parseExcel")
    public void parseExcel() throws Exception {
        FileInputStream inputStream = new FileInputStream(BASE_DIRECTORY_PATH + "HSSF测试method.xls");
        HSSFWorkbook workbook = new HSSFWorkbook(inputStream);
        HSSFSheet sheet = workbook.getSheetAt(0);
        FormulaEvaluator evaluator = new HSSFFormulaEvaluator(workbook);
        for (Row cells : sheet) {
            for (Cell cell : cells) {
                Object cellValueType = excelUtil.getObjectCellValue(cell);
                System.out.println("cellValueType = " + cellValueType);
                if (cellValueType instanceof Date)
                    System.out.println("cellValueType After = " + new SimpleDateFormat("yyyy-mm-dd").format(cellValueType));
                String cellValue = excelUtil.getStringCellValue(cell, evaluator);
                System.out.println("cellValue = " + cellValue);
            }
        }

        excelUtil.closeStream(workbook, inputStream);

    }

    /**
     * 条件格式化、ConditionalFormatting
     *
     * @param sheet 工作簿
     */
    public void formating(Sheet sheet) {
        SheetConditionalFormatting conditionalFormatting = sheet.getSheetConditionalFormatting();

        ConditionalFormattingRule rule1 = conditionalFormatting.createConditionalFormattingRule(ComparisonOperator.EQUAL, "0");
        ConditionalFormattingRule rule2 = conditionalFormatting.createConditionalFormattingRule(ComparisonOperator.BETWEEN, "0", "10");
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
        //设置自动根据内容扩展列宽
        sheet.autoSizeColumn(3);

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

        //添加日期格式化的样式、(两种数据格式化方式)
        HSSFDataFormat dataFormat = workbook.createDataFormat();
        style.setDataFormat(dataFormat.getFormat("yyyy-mm-dd h:mm:ss"));
//        style.setDataFormat(HSSFDataFormat.getBuiltinFormat("yyyy-mm-dd"));

        //TODO: 可以通过打印或者源码查看 默认的格式、
        List<String> builtinFormats = HSSFDataFormat.getBuiltinFormats();
        builtinFormats.forEach(System.out::println);

        HSSFRow row = sheet.createRow(0);
        HSSFRow row1 = sheet.createRow(1);
        CellUtil.createCell(row, 0, "CellUtil--00", style);
        CellUtil.createCell(row, 1, "CellUtil--01", style);
        CellUtil.createCell(row1, 0, "CellUtil--10", style);
        CellUtil.createCell(row1, 1, "CellUtil--11", style);


        double excelDate = HSSFDateUtil.getExcelDate(new Date());
        Date javaDate = HSSFDateUtil.getJavaDate(excelDate);

        System.out.println("excelDate = " + excelDate);
        System.out.println("javaDate = " + javaDate);

        HSSFCell cell = row.createCell(3);
        cell.setCellStyle(style);//此单元格应用了日期格式化的样式、
        cell.setCellValue(new Date());
        HSSFCell cell4 = row.createCell(4);
        cell4.setCellValue(excelDate);//日期需要格式化、否则无法正常显示、

        FileOutputStream outputStream = new FileOutputStream(BASE_DIRECTORY_PATH + "HSSF测试utils.xls");
        workbook.write(outputStream);

        excelUtil.closeStream(workbook, outputStream);
    }

    /**
     * 为单元格添加计算公式、
     */
    @GetMapping("formula")
    public void formula() throws Exception {
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet();
        HSSFRow row = sheet.createRow(0);
        HSSFCell cell = row.createCell(0);
        cell.setCellFormula("2+3*4");//设置公式
        cell = row.createCell(1);
        cell.setCellValue(10);
        cell = row.createCell(2);
        cell.setCellFormula("A1*B1");//设置公式 "sum(A1,C1)" "sum(B1:D1)" 等Excel函数都可以使用、

        FileOutputStream outputStream = new FileOutputStream(BASE_DIRECTORY_PATH + "HSSF测试formula.xls");
        workbook.write(outputStream);

        excelUtil.closeStream(workbook, outputStream);
    }

    /**
     * 画简单的形状、常用来插入图片、
     *
     * @throws Exception ex
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
        richTextString.applyFont(excelUtil.getFont(workbook, IndexedColors.RED.index, true, null));
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
     * @throws Exception ex
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
        HSSFPicture picture = patriarch.createPicture(anchorPicture, pictureIndex);

        //设置图片
        picture.resize();//自动调节图片大小,图片位置信息可能丢失
        read.flush();

        FileOutputStream outputStream = new FileOutputStream(BASE_DIRECTORY_PATH + "HSSF测试pictures.xls");
        workbook.write(outputStream);

        excelUtil.closeStream(workbook, outputStream);
    }

    /**
     * 插入多张图片、
     *
     * @throws Exception ex
     */
    @GetMapping("addPictures")
    public void addPictures() throws Exception {
        FileInputStream inputStream = new FileInputStream(XLS_TEMPLATE_FILE_PATH);
        HSSFWorkbook workbook = new HSSFWorkbook(inputStream);
        HSSFSheet sheet = workbook.getSheetAt(0);
        HSSFRow row5 = sheet.getRow(4);
        row5.setHeightInPoints(sheet.getDefaultRowHeightInPoints());

        BufferedImage image1 = ImageIO.read(new File(BASE_DIRECTORY_PATH + "1.jpg"));
        BufferedImage image2 = ImageIO.read(new File(BASE_DIRECTORY_PATH + "2.jpg"));
        BufferedImage image3 = ImageIO.read(new File(BASE_DIRECTORY_PATH + "3.jpg"));
        BufferedImage image4 = ImageIO.read(new File(BASE_DIRECTORY_PATH + "4.jpg"));
        ByteArrayOutputStream byteArrayOS1 = new ByteArrayOutputStream();
        ByteArrayOutputStream byteArrayOS2 = new ByteArrayOutputStream();
        ByteArrayOutputStream byteArrayOS3 = new ByteArrayOutputStream();
        ByteArrayOutputStream byteArrayOS4 = new ByteArrayOutputStream();
        ImageIO.write(image1, "JPG", byteArrayOS1);
        ImageIO.write(image2, "JPG", byteArrayOS2);
        ImageIO.write(image3, "JPG", byteArrayOS3);
        ImageIO.write(image4, "JPG", byteArrayOS4);

        int index1 = workbook.addPicture(byteArrayOS1.toByteArray(), Workbook.PICTURE_TYPE_JPEG);
        int index2 = workbook.addPicture(byteArrayOS2.toByteArray(), Workbook.PICTURE_TYPE_JPEG);
        int index3 = workbook.addPicture(byteArrayOS3.toByteArray(), Workbook.PICTURE_TYPE_JPEG);
        int index4 = workbook.addPicture(byteArrayOS4.toByteArray(), Workbook.PICTURE_TYPE_JPEG);

        log.info("index1 = {}, index2 = {}, index3 = {}, index4 = {}", index1, index2, index3, index4);

        HSSFPatriarch patriarch = sheet.createDrawingPatriarch();
        /*
         *   0<= dx <=1023
         *   0<= dy <=255
         *   锚点的起点：左上角(dx1,dy1)、
         *   锚点的终点：右上角(dx2,dy2)、
         * */
        HSSFClientAnchor anchor1 = new HSSFClientAnchor(16, 48, 975, 239, (short) 5, 1, (short) 5, 1);//设置在同一个单元格内、并有间距
        HSSFClientAnchor anchor2 = new HSSFClientAnchor(16, 48, 975, 239, (short) 5, 2, (short) 5, 2);
        HSSFClientAnchor anchor3 = new HSSFClientAnchor(16, 48, 975, 239, (short) 5, 3, (short) 5, 3);
        HSSFClientAnchor anchor4 = new HSSFClientAnchor(16, 48, 975, 239, (short) 5, 4, (short) 5, 4);
        patriarch.createPicture(anchor1, index1);
        patriarch.createPicture(anchor2, index2);
        patriarch.createPicture(anchor3, index3);
        patriarch.createPicture(anchor4, index4);

        FileOutputStream outputStream = new FileOutputStream(XLS_TEMPLATE_FILE_PATH);
        workbook.write(outputStream);

        image1.flush();
        image2.flush();
        image3.flush();
        image4.flush();

        excelUtil.closeStream(workbook, outputStream);
    }

    /**
     * 利用HSSFPatriarch对象,创建批注、注释.
     */
    @GetMapping("comment")
    public void comment() throws Exception {
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet();
        HSSFPatriarch patriarch = sheet.createDrawingPatriarch();
        HSSFClientAnchor anchor = patriarch.createAnchor(0, 0, 0, 0, 5, 1, 8, 3);//创建批注位置(锚)
        HSSFComment comment = patriarch.createCellComment(anchor);//创建批注
        comment.setString(new HSSFRichTextString("这是一个批注段落！"));//设置批注内容
        comment.setAuthor("小66");//设置批注作者
//        comment.setFillColor(IndexedColors.RED.index);
//        comment.setBackgroundImage();//背景图片
        comment.setVisible(false);//设置批注 默认显示-->true
        HSSFCell cell = sheet.createRow(2).createCell(1);
        cell.setCellValue("新建批注");
        cell.setCellComment(comment);//把批注赋值给单元格

        FileOutputStream outputStream = new FileOutputStream(BASE_DIRECTORY_PATH + "HSSF测试comment.xls");
        workbook.write(outputStream);

        excelUtil.closeStream(workbook, outputStream);
    }

    /**
     * 超链接引入图片、
     */
    @GetMapping("hyperlink")
    public void hyperlink() throws Exception {
        HSSFWorkbook workbook = new HSSFWorkbook();
        /*
         *   HyperlinkType.URL：关联到指定Url、
         *   HyperlinkType.FILE：关联到目录文件、
         *   HyperlinkType.DOCUMENT：工作簿中的位置、
         * */
        HSSFHyperlink hyperlink = workbook.getCreationHelper().createHyperlink(HyperlinkType.URL);
        hyperlink.setShortFilename("驴子");
        hyperlink.setAddress("http://p20.qhimgs3.com/dr/240_240_/t011c15f8c12e731c01.jpg?t=1558300397");
        HSSFSheet sheet = workbook.createSheet();
        HSSFCell cell = sheet.createRow(0).createCell(0);
        HSSFCellStyle style = workbook.createCellStyle();
        Font font = excelUtil.getFont(workbook, IndexedColors.BLUE.index, false, null);
        style.setFont(font);
        cell.setCellStyle(style);//为单元格设置样式,否则样式不生效、
        cell.setHyperlink(hyperlink);
        cell.setCellValue("hyperlink");

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

    /**
     * 打印设置、PrintSetup
     */
    @GetMapping("printSet")
    public void printSet() {
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet("Test0");// 创建工作表(Sheet)
        HSSFPrintSetup print = sheet.getPrintSetup();//得到打印对象
        print.setLandscape(false);//true，则表示页面方向为横向；否则为纵向
        print.setScale((short) 80);//缩放比例80%(设置为0-100之间的值)
        print.setFitWidth((short) 2);//设置页宽
        print.setFitHeight((short) 4);//设置页高
        print.setPaperSize(HSSFPrintSetup.A4_PAPERSIZE);//纸张设置
        print.setUsePage(true);//设置打印起始页码不使用"自动"
        print.setPageStart((short) 6);//设置打印起始页码
        sheet.setPrintGridlines(true);//设置打印网格线
        print.setNoColor(true);//值为true时，表示单色打印
        print.setDraft(true);//值为true时，表示用草稿品质打印
        print.setLeftToRight(true);//true表示“先行后列”；false表示“先列后行”
        print.setNotes(true);//设置打印批注
        sheet.setAutobreaks(false);//Sheet页自适应页面大小

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

        //设置填充属性(背景和纹理):-->不常用、
//        cellStyle.setFillForegroundColor(HSSFColor.GREEN.index);//设置图案颜色
//        cellStyle.setFillBackgroundColor(HSSFColor.RED.index);//设置图案背景色
//        cellStyle.setFillPattern(FillPatternType.BIG_SPOTS);//设置图案样式

        //设置字体属性
        cellStyle.setFont(font);

        //设置换行
        cellStyle.setWrapText(true);

        //设置锁定、
        cellStyle.setLocked(true);

        cell.setCellStyle(cellStyle);
    }

    /**
     * 遍历Excel表格、
     * row.getCell(0, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);//指定当前单元格为null and blank cells时的策略、(暂时不知道有何用)
     *
     * @param workbook Excel文件
     */
    public void simpleIterator(Workbook workbook, List<Map<String, Object>> dataList, Map<String, PictureData> pictureDataMap) {
        DataFormatter formatter = new DataFormatter();
        for (Sheet sheet : workbook) {//sheetIterator
            Row row0 = sheet.getRow(0);
            for (Row row : sheet) {//rowIterator
                Map<String, Object> map = new HashMap<>();
                for (Cell cell : row) {//cellIterator
                    CellReference cellRef = new CellReference(row.getRowNum(), cell.getColumnIndex());
                    log.info("cellRef.formatAsString = {}", cellRef.formatAsString());
                    String formatValue = formatter.formatCellValue(cell);//不进行formatter、则需要对单元格内容进行判断后再获取、
                    CellAddress address = cell.getAddress();//获取当前单元格的坐标
                    log.info("row = {}, column = {}, value = {}", address.getRow(), address.getColumn(), formatValue);
                    if (cell.getHyperlink() != null) {
                        String linkAddr = cell.getHyperlink().getAddress();
                        map.put(row0.getCell(address.getColumn()).getStringCellValue(), formatValue + ":" + linkAddr);//key是表头、
                        continue;
                    }
//                    cell.getCellComment();//获取标注、注释
                    map.put(row0.getCell(address.getColumn()).getStringCellValue(), formatValue);//key是表头、
                }
                pictureDataMap.keySet().stream().filter(key -> key.startsWith(row.getRowNum() + "-"))
                        .forEach(key -> map.put(row0.getCell(Integer.parseInt(key.split("-")[1])).getStringCellValue(), pictureDataMap.get(key)));
                dataList.add(map);
            }
        }
    }

    @GetMapping("parseComplexXls")
    public void parseComplexXls() throws Exception {
        HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(XLS_TEMPLATE_FILE_PATH));
        HSSFSheet sheetAt = workbook.getSheetAt(0);
//        List<Map<String, Object>> mapList = excelUtil.parseSimpleExcel(sheetAt,null);
        List<Map<String, Object>> mapList = excelUtil.parseComplexXls(sheetAt, workbook.getCreationHelper().createFormulaEvaluator());

        for (Map<String, Object> map : mapList) {
            //这个Key是图片列的表头,如果图片不是一列、则对Map进行遍历.
            Object pictureData = map.get("图片");
            if (pictureData instanceof PictureData) {
                ListenableFuture<String> future = excelUtil.writeImage(BASE_DIRECTORY_PATH, (PictureData) pictureData);
                /*
                 *   ListenableFuture可以直接处理异步返回结果 --> 不需要再使用Executors.newCachedThreadPool().execute()、
                 *   Future没有办法直接处理异步返回结果、
                 * */
                future.addCallback(result -> {
                    try {
                        String path = future.get();
                        map.put("图片", path);
                        //执行插入数据的操作、
                        System.out.println("成功插入数据: " + map);
                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                }, ex -> System.out.println("抛出异常 = " + ex.getLocalizedMessage()));

//                Executors.newCachedThreadPool().execute(() -> {
//                    while (!future.isDone()) {
//                        try {
//                            TimeUnit.MILLISECONDS.sleep(1500);
//                        } catch (InterruptedException e) {
//                            e.printStackTrace();
//                        }
//                    }
//                    try {
//                        String path = future.get();
//                        map.put("图片", path);
//                        //执行插入数据的操作、
//                        System.out.println("成功插入数据: " + map);
//                    } catch (Exception e) {
//                        e.printStackTrace();
//                    }
//                });
            }
        }
        System.out.println("mapList = " + mapList.size());
        workbook.close();
    }

    @GetMapping("addData")
    public void addData() throws Exception {
        List<Map<String, Object>> mapList = new ArrayList<>();
        Map<String, Object> data = new HashMap<>();
        data.put("序号", 1);
        data.put("文字", "中国");
        data.put("日期", LocalDate.now());
        data.put("数字", 199.99);
        data.put("超链接", "百度---http://www.baidu.com");
        data.put("图片", BASE_DIRECTORY_PATH + "1.jpg");

        Map<String, Object> data2 = new HashMap<>();
        data2.put("序号", 2);
        data2.put("文字", "中国");
        data2.put("日期", LocalDate.now());
        data2.put("数字", 199.99);
        data2.put("超链接", "百度---http://www.baidu.com");
        data2.put("图片", BASE_DIRECTORY_PATH + "2.jpg");

        Map<String, Object> data3 = new HashMap<>();
        data3.put("序号", 3);
        data3.put("文字", "中国");
        data3.put("日期", LocalDateTime.now());
        data3.put("数字", 199.99);
        data3.put("超链接", "百度---http://www.baidu.com");
        data3.put("图片", BASE_DIRECTORY_PATH + "3.jpg");

        mapList.add(data);
        mapList.add(data2);
        mapList.add(data3);

        FileInputStream fileInputStream = new FileInputStream(XLS_TEMPLATE_FILE_PATH);
        HSSFWorkbook workbook = new HSSFWorkbook(fileInputStream);
        HSSFSheet sheetAt = workbook.getSheetAt(0);
        sheetAt.autoSizeColumn(2);

        HSSFRow lastRow = sheetAt.getRow(sheetAt.getLastRowNum());
        Font linkFont = excelUtil.getFont(workbook, IndexedColors.BLUE.index, false, Font.U_SINGLE);
        float lastRowHeightInPoints = lastRow.getHeightInPoints();
        HSSFCellStyle lastRowStyle = lastRow.getRowStyle();

        HSSFPatriarch patriarch = sheetAt.createDrawingPatriarch();
        CellStyle newStyle = excelUtil.getStyle(workbook);

        CellStyle linkStyle = excelUtil.getStyle(workbook, linkFont, HorizontalAlignment.LEFT);
        HSSFRow row0 = sheetAt.getRow(sheetAt.getFirstRowNum());
        System.out.println("lastRow.getRowNum() = " + lastRow.getRowNum());
        System.out.println("lastRowStyle = " + lastRowStyle);
        System.out.println("lastRowHeightInPoints = " + lastRowHeightInPoints);
        System.out.println("sheetAt.getDefaultRowHeightInPoints() = " + sheetAt.getDefaultRowHeightInPoints());


        if (row0 == null) return;
        int lastRowNum = sheetAt.getLastRowNum();
        for (Map<String, Object> map : mapList) {
            lastRowNum++;
            HSSFRow row = sheetAt.createRow(lastRowNum);
            row.setHeightInPoints(lastRowHeightInPoints);
            //复制最后一行(原始)的RowStyle、
            if (!excelUtil.isRowEmpty(lastRow) && lastRowStyle != null) {
                row.setRowStyle(lastRowStyle);
            }//部分样式无法正常复制、
            for (int column = 0; column < map.entrySet().size(); column++) {
                //表头、
                String key = row0.getCell(column).getStringCellValue();
                Object value = map.get(key);
                //图片路径、
                if (value.toString().startsWith(BASE_DIRECTORY_PATH)) {
                    String absFilePath = value.toString();
                    createPicture(workbook, patriarch, row, (short) column, absFilePath);
                } else if (value.toString().contains("---")) {
                    String[] split = value.toString().split("---");//超链接需要指定固定的格式、
                    HSSFCell cell = row.createCell(column);
                    cell.setCellStyle(linkStyle);//超链接样式无法复制下划线,需要单独添加
                    HSSFHyperlink hyperlink = workbook.getCreationHelper().createHyperlink(HyperlinkType.URL);
                    hyperlink.setAddress(split[1]);
                    //没生效、如果不进行cell.setCellValue(split[0]);则显示空白
                    hyperlink.setShortFilename(split[0]);
                    cell.setHyperlink(hyperlink);
                    cell.setCellValue(split[0]);

                } else {
                    HSSFCell cell = row.createCell(column);
                    if (excelUtil.isRowEmpty(lastRow) || lastRowStyle == null)
                        cell.setCellStyle(newStyle);
                    excelUtil.setCellValue(cell, value);
                }
            }
        }

        workbook.write(new File(XLS_TEMPLATE_FILE_PATH));
        excelUtil.closeStream(workbook, fileInputStream);
    }

    private void createPicture(HSSFWorkbook workbook, HSSFPatriarch patriarch, HSSFRow row, short i, String absFilePath) throws IOException {
        ByteArrayOutputStream byteArrayOS = excelUtil.readImage(absFilePath);
        //获取MimeType、
        String contentType = Files.probeContentType(Paths.get(absFilePath));

        int index;
        switch (contentType) {
            case MimeTypeUtils.IMAGE_JPEG_VALUE:
                index = workbook.addPicture(byteArrayOS.toByteArray(), Workbook.PICTURE_TYPE_JPEG);
                break;
            case MimeTypeUtils.IMAGE_PNG_VALUE:
                index = workbook.addPicture(byteArrayOS.toByteArray(), Workbook.PICTURE_TYPE_PNG);
                break;
            default:
                index = -1;
        }
        if (index != -1) {
            //图片所占单元格大小、按需设置、
            HSSFClientAnchor anchor = excelUtil.getAnchor(16, 48, 975, 239, i, row.getRowNum(), i, row.getRowNum());
            patriarch.createPicture(anchor, index);
        }
    }

    @GetMapping("addDataByUtils")
    public void addDataByUtils() throws Exception {
        List<Map<String, Object>> mapList = new ArrayList<>();
        Map<String, Object> data = new HashMap<>();
        data.put("序号", 1);
        data.put("文字", "中国");
        data.put("日期", LocalDate.now());
        data.put("数字", 199.99);
        data.put("超链接", "百度---http://www.baidu.com");
        data.put("图片", BASE_DIRECTORY_PATH + "1.jpg");

        Map<String, Object> data2 = new HashMap<>();
        data2.put("序号", 2);
        data2.put("文字", "中国");
        data2.put("日期", new Date());
        data2.put("数字", 199.99);
        data2.put("超链接", "百度---http://www.baidu.com");
        data2.put("图片", BASE_DIRECTORY_PATH + "2.jpg");

        Map<String, Object> data3 = new HashMap<>();
        data3.put("序号", 3);
        data3.put("文字", "中国");
        data3.put("日期", Calendar.getInstance());//这个格式无法正常解析、
        data3.put("数字", 199.99);
        data3.put("超链接", "百度---http://www.baidu.com");
        data3.put("图片", BASE_DIRECTORY_PATH + "3.jpg");

        mapList.add(data);
        mapList.add(data2);
        mapList.add(data3);

        excelUtil.addNonEmptyRow(XLS_TEMPLATE_FILE_PATH, BASE_DIRECTORY_PATH, mapList, null);
        System.out.println("测试异步新增：" + mapList.size() + "条数据");
    }

    @GetMapping("createXlsWithData")
    public void createXlsWithData() throws Exception {
        List<Map<String, Object>> mapList = new ArrayList<>();
        Map<String, Object> data = new HashMap<>();
        data.put("序号", 1);
        data.put("文字", "中国");
        data.put("日期", LocalDate.now());
        data.put("数字", 199.99);
        data.put("超链接", "百度---http://www.baidu.com");
        data.put("图片", BASE_DIRECTORY_PATH + "1.jpg");

        Map<String, Object> data2 = new HashMap<>();
        data2.put("序号", 2);
        data2.put("文字", "中国");
        data2.put("日期", LocalDate.now());
        data2.put("数字", 199.99);
        data2.put("超链接", "百度---http://www.baidu.com");
        data2.put("图片", BASE_DIRECTORY_PATH + "2.jpg");

        Map<String, Object> data3 = new HashMap<>();
        data3.put("序号", 3);
        data3.put("文字", "中国");
        data3.put("日期", LocalDateTime.now());
        data3.put("数字", 199.99);
        data3.put("超链接", "百度---http://www.baidu.com");
        data3.put("图片", BASE_DIRECTORY_PATH + "3.jpg");

        mapList.add(data);
        mapList.add(data2);
        mapList.add(data3);

        List<String> headerList = new ArrayList<>();
        headerList.add("序号");
        headerList.add("文字");
        headerList.add("日期");
        headerList.add("数字");
        headerList.add("超链接");
        headerList.add("图片");

        excelUtil.createXlsAndInsertData(BASE_DIRECTORY_PATH + "HSSF测试createWithData.xls", BASE_DIRECTORY_PATH, headerList,
                mapList, 28F, 11, null);
        System.out.println("测试异步创建Excel文件,并填充：" + mapList.size() + "条数据");
    }
}
