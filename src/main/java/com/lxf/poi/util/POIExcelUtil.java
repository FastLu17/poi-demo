package com.lxf.poi.util;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.springframework.stereotype.Component;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * 操作Excel的工具类、
 *
 * @author 小66
 * @create 2019-08-12 16:29
 **/
@Component
@Slf4j
public class POIExcelUtil {
    private final String BASE_URL = "C:\\Users\\Administrator\\Desktop\\POI\\";

    /**
     * 读入excel的内容转换成字符串、
     * Date：格式化为:"yyyy-MM-dd hh:mm:ss"
     * Numeric：格式化为:"#.##" -->两位小数、
     *
     * @param cell 单元格
     * @return StringValue
     */
    public String getStringCellValue(Cell cell) {
        SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd hh:mm:ss");
        DecimalFormat decimalFormat = new DecimalFormat("#.##");//两位小数
        String cellValue = "";
        if (cell == null)
            return cellValue;

        //TODO:3.5版本没有不过期的方法、getCellType()也是过期的方法、
        switch (cell.getCellTypeEnum()) {
            case STRING:
                cellValue = cell.getStringCellValue();
                break;
            case BLANK:
                cellValue = "";
                break;
            case ERROR:
                cellValue = "";
                break;
            case FORMULA:
                cellValue = cell.getCellFormula();
                break;
            case BOOLEAN:
                cellValue = String.valueOf(cell.getBooleanCellValue());
                break;
            case NUMERIC:
                if (HSSFDateUtil.isCellDateFormatted(cell)) {
                    Date date = cell.getDateCellValue();
                    cellValue = dateFormat.format(date);
                } else {
                    cellValue = decimalFormat.format((cell.getNumericCellValue()));
                }
                break;
        }
        return cellValue;
    }

    /**
     * 获取Cell的原始类型值、读取是什么类型,就是什么类型、
     * Date、byte(Error)、boolean、String、double、RichTextString
     *
     * @param cell 单元格
     * @return cell的value值、类型不定、
     */
    public Object getObjectCellValue(Cell cell) {
        if (cell == null)
            return null;

        switch (cell.getCellTypeEnum()) {
            case STRING:
                return cell.getStringCellValue();
            case BLANK:
                return null;
            case ERROR:
                return cell.getErrorCellValue();
            case FORMULA:
                return cell.getStringCellValue();
            case BOOLEAN:
                return cell.getBooleanCellValue();
            case NUMERIC:
                if (HSSFDateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue();
                } else {
                    return cell.getNumericCellValue();
                }
            default:
                return null;
        }
    }


    /**
     * 中等(MEDIUM)的边框、
     *
     * @param workbook    xls文档、
     * @param font        字体, 默认：13字号、Consolas、黑色
     * @param borderStyle 边框样式, 默认：BorderStyle.THIN
     * @param borderColor 边框颜色, 默认：IndexedColors.BLACK.index
     * @param wrapText    是否换行, 默认：false
     * @param lock        是否锁定, 默认：false
     * @return CellStyle
     */
    public CellStyle getStyle(Workbook workbook, Font font, BorderStyle borderStyle, short borderColor, boolean wrapText, boolean lock) {
        CellStyle style = workbook.createCellStyle();

        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);

        //设置边框属性
        style.setBorderBottom(borderStyle);
        style.setBottomBorderColor(borderColor);
        style.setBorderLeft(borderStyle);
        style.setLeftBorderColor(borderColor);
        style.setBorderRight(borderStyle);
        style.setRightBorderColor(borderColor);
        style.setBorderTop(borderStyle);
        style.setTopBorderColor(borderColor);

        style.setFont(font);//字体
        style.setWrapText(wrapText);//自动换行

        //设置填充属性(背景和纹理):-->不常用
        /*
         * style.setFillForegroundColor(IndexedColors.GREEN.index);//设置图案颜色
         * style.setFillBackgroundColor(HSSFColor.RED.index);//设置图案背景色
         * style.setFillPattern(FillPatternType.BIG_SPOTS);//设置图案样式
         * */
        //设置锁定、
        style.setLocked(lock);

        return style;
    }

    /**
     * 获取自定义的默认样式、黑色-细的(THIN)边框,不换行,不锁定,自定义默认字体(Consolas)、
     *
     * @param workbook workbook
     * @return CellStyle
     */
    public CellStyle getStyle(Workbook workbook) {
        return getStyle(workbook, getFont(workbook), BorderStyle.THIN,
                IndexedColors.BLACK.index, false, false);
    }

    /**
     * 不要循环创建字体样式,尽量重用、
     *
     * @param workbook
     * @param fontHeightInPoints 字号大小。 200 FontHeight <==> 10 FontHeightInPoints
     * @param fontName           字体名称
     * @param color              eg:IndexedColors.GREEN.index,HSSFColor.RED.index
     * @param bold               是否加粗
     * @param italic             是否设置斜体
     * @param strikeout          是否设置删除线
     * @param underLine          下划线、 eg；Font.U_SINGLE
     * @return Font
     */
    public Font getFont(Workbook workbook, int fontHeightInPoints, String fontName,
                        short color, boolean bold, boolean italic, boolean strikeout, byte underLine) {
        Font font = workbook.createFont();
        font.setFontHeightInPoints((short) fontHeightInPoints);//设置字号、
        font.setFontName(fontName);
        font.setBold(bold);
        font.setItalic(italic);//斜体
        font.setStrikeout(strikeout);//删除线
        font.setUnderline(underLine);

        //设置字体颜色
        font.setColor(color);

        return font;
    }

    /**
     * 获取自定义的默认样式、默认：13字号、Consolas、黑色、不加粗、没有下划线、不斜体
     * 注意：不要循环创建字体样式,尽量重用、
     *
     * @param workbook workbook
     * @return Font
     */
    public Font getFont(Workbook workbook) {
        return getFont(workbook, 13, "Consolas",
                IndexedColors.BLACK.index, false, false, false, Font.U_NONE);
    }

    /**
     * 获取自定义的默认样式、默认：13字号、Consolas、没有下划线、不斜体
     * 注意：不要循环创建字体样式,尽量重用、
     *
     * @param workbook workbook
     * @param color    eg:IndexedColors.GREEN.index,HSSFColor.RED.index
     * @param bold     是否加粗
     * @return Font
     */
    public Font getFont(Workbook workbook, short color, boolean bold) {
        return getFont(workbook, 13, "Consolas",
                color, bold, false, false, Font.U_NONE);
    }


    /**
     * 不安全的合并单元格、
     *
     * @param sheet        工作簿
     * @param rangeAddress 合并的范围、
     */
    public void mergingCellsUnsafe(Sheet sheet, CellRangeAddress rangeAddress) {
        sheet.addMergedRegionUnsafe(rangeAddress);//两个单元格都有值,都会保留,但是只显示一个单元格的值、
    }

    /**
     * 安全的合并单元格、
     *
     * @param sheet        工作簿
     * @param rangeAddress 合并的范围、
     */
    public void mergingCells(Sheet sheet, CellRangeAddress rangeAddress) {
        sheet.addMergedRegion(rangeAddress);//没测试过具体效果、
    }

    /**
     * 简单的遍历Excel表格中指定的sheet、适用于没有图片的表格、
     *
     * @param sheet 工作簿
     */
    public List<Map<String, Object>> simpleIterator(Sheet sheet) {
        List<Map<String, Object>> dataList = new ArrayList<>();
        DataFormatter formatter = new DataFormatter();
        Row row0 = sheet.getRow(0);
        for (Row row : sheet) {//rowIterator
            if (row.getRowNum() == 0)
                continue;
            Map<String, Object> map = new HashMap<>();
            for (Cell cell : row) {//cellIterator
                CellReference cellRef = new CellReference(row.getRowNum(), cell.getColumnIndex());
                log.info("cellRef.formatAsString = {}", cellRef.formatAsString());
                String formatValue = formatter.formatCellValue(cell);//不进行formatter、则需要对单元格内容进行判断后再获取、
                CellAddress address = cell.getAddress();//获取当前单元格的坐标
                if (cell.getHyperlink() != null) {
                    String linkAddr = cell.getHyperlink().getAddress();
                    map.put(row0.getCell(address.getColumn()).getStringCellValue(), formatValue + ":" + linkAddr);//key是表头、
                    continue;
                }
//                    cell.getCellComment();//获取标注、注释
                map.put(row0.getCell(address.getColumn()).getStringCellValue(), formatValue);//key是表头、
            }
            dataList.add(map);
        }
        return dataList;
    }

    public void closeStream(Workbook workbook, FileOutputStream outputStream) throws IOException {
        if (outputStream != null)
            outputStream.close();
        if (workbook != null)
            workbook.close();
    }

    public void closeStream(Workbook workbook, FileInputStream inputStream) throws IOException {
        if (workbook != null) {
            workbook.close();
        }
        if (inputStream != null)
            inputStream.close();
    }

    public void closeStream(Workbook workbook, FileOutputStream outputStream, FileInputStream inputStream) throws IOException {
        if (outputStream != null)
            outputStream.close();
        if (workbook != null)
            workbook.close();
        if (inputStream != null)
            inputStream.close();
    }
}
