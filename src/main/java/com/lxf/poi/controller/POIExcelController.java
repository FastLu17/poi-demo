package com.lxf.poi.controller;

import com.lxf.poi.util.POIExcelUtil;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.extractor.ExcelExtractor;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

/**
 * @author 小66
 * @create 2019-08-12 21:43
 **/
@RestController
@Slf4j
public class POIExcelController {

    @Autowired
    private POIExcelUtil excelUtil;

    private final String BASE_FILE_PATH = "C:\\Users\\Administrator\\Desktop\\POI\\";

    @GetMapping("createXls")
    public String createXls() throws Exception {
//        POIFSFileSystem system = new POIFSFileSystem(new File(BASE_FILE_PATH + "HSSF测试read.xls"));
//        Workbook book = WorkbookFactory.create(system);
//        HSSFWorkbook workbook = new HSSFWorkbook(system);  //多种方式可以创建Workbook对象、
        HSSFWorkbook workbook = new HSSFWorkbook();
        Sheet sheet = workbook.createSheet();
        Row row = sheet.createRow(1);
        row.setHeightInPoints(30);

        createCell(workbook, row, 0, HorizontalAlignment.CENTER, VerticalAlignment.BOTTOM, "Align It");
        createCell(workbook, row, 1, HorizontalAlignment.CENTER_SELECTION, VerticalAlignment.BOTTOM, "Align It");
        createCell(workbook, row, 2, HorizontalAlignment.FILL, VerticalAlignment.CENTER, "Align It");
        createCell(workbook, row, 3, HorizontalAlignment.GENERAL, VerticalAlignment.CENTER, "Align It");
        createCell(workbook, row, 4, HorizontalAlignment.JUSTIFY, VerticalAlignment.JUSTIFY, "Align It");
        createCell(workbook, row, 5, HorizontalAlignment.LEFT, VerticalAlignment.TOP, "Align It");
        createCell(workbook, row, 6, HorizontalAlignment.RIGHT, VerticalAlignment.TOP, "Align It");

        FileOutputStream outputStream = new FileOutputStream(BASE_FILE_PATH + "HSSF测试create.xls");
        workbook.write(outputStream);

        excelUtil.closeStream(workbook, outputStream);

        return BASE_FILE_PATH + "HSSF测试create.xls";
    }

    @GetMapping("readXls")
    public String readXls() throws Exception {
        POIFSFileSystem system = new POIFSFileSystem(new File(BASE_FILE_PATH + "HSSF测试read.xls"));
        Workbook book = WorkbookFactory.create(system);
//        simpleIterator(book);
        simpleIterator(book);
        book.close();
        return "...";
    }

    @GetMapping("getXlsText")
    public String getXlsText() throws Exception {
        POIFSFileSystem system = new POIFSFileSystem(new File(BASE_FILE_PATH + "HSSF测试read.xls"));
        HSSFWorkbook workbook = new HSSFWorkbook(system);
        //Text Extraction：文字抽取
        ExcelExtractor extractor = new ExcelExtractor(workbook);

        extractor.setFormulasNotResults(false);//设置公式有无返回值--> true：输出公式、 false：输出结果
        extractor.setIncludeSheetNames(false);
        return extractor.getText();
    }

    @GetMapping("mergingCells")
    public void createMergingCells() throws Exception {
        mergingCells(new CellRangeAddress(1, 1, 0, 1), BASE_FILE_PATH + "HSSF测试read.xls");

    }

    private void createCell(Workbook wb, Row row, int column, HorizontalAlignment halign, VerticalAlignment valign, String value) {
        Cell cell = row.createCell(column);
        cell.setCellValue(value);

        //指定样式、
        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setAlignment(halign);
        cellStyle.setVerticalAlignment(valign);

        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBottomBorderColor(IndexedColors.YELLOW.getIndex());
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setLeftBorderColor(IndexedColors.GREEN.getIndex());
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setRightBorderColor(IndexedColors.BLUE.getIndex());
        cellStyle.setBorderTop(BorderStyle.MEDIUM_DASHED);
        cellStyle.setTopBorderColor(IndexedColors.RED.getIndex());

        //设置填充属性
        cellStyle.setFillBackgroundColor(IndexedColors.AQUA.getIndex());
        cellStyle.setFillPattern(FillPatternType.BIG_SPOTS);

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
                    String formatValue = formatter.formatCellValue(cell);//不进行formatter、则需要对单元格内容进行判断后再获取、
                    CellAddress address = cell.getAddress();//获取当前单元格的坐标
                    log.info("row = {}, column = {}, value = {}", address.getRow(), address.getColumn(), cell.getStringCellValue());
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
}
