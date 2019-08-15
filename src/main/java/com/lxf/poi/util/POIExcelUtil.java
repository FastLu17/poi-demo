package com.lxf.poi.util;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.springframework.lang.NonNull;
import org.springframework.lang.Nullable;
import org.springframework.scheduling.annotation.Async;
import org.springframework.stereotype.Component;
import org.springframework.util.StringUtils;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.*;
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
    private final DataFormatter formatter = new HSSFDataFormatter();
    private final String BASE_URL = "C:\\Users\\Administrator\\Desktop\\POI\\";

    /**
     * 读入excel的内容转换成字符串,存在公式时,获取对应公式、
     * Date：格式化为:"yyyy-MM-dd hh:mm:ss"
     *
     * @param cell 单元格
     * @return StringValue
     */
    public String getStringCellValue(Cell cell) {
        return getStringCellValue(cell, null);
    }

    /**
     * 读入excel的内容转换成字符串、存在公式时,获取计算后再获取值、
     * Date：格式化为:"yyyy-MM-dd hh:mm:ss"
     *
     * @param cell      单元格
     * @param evaluator 公式计算器
     * @return StringValue
     */
    public String getStringCellValue(Cell cell, @Nullable FormulaEvaluator evaluator) {
        /*
         * 类似此操作、只是定义了对日期类型的不同格式化、
         * DataFormatter formatter = new HSSFDataFormatter();
         * formatter.formatCellValue(cell);
         * */
        //hh是12小时制的时间、
        SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        if (cell == null)
            return "";
        //TODO:3.5版本没有不过期的方法、getCellType()也是过期的方法、
        CellType cellTypeEnum = cell.getCellTypeEnum();
        if (cellTypeEnum == CellType.FORMULA) {
            if (evaluator == null) {
                return cell.getCellFormula().trim();
            }
            cellTypeEnum = evaluator.evaluate(cell).getCellTypeEnum();//计算公式后的结果类型、
        }
        switch (cellTypeEnum) {
            case STRING:
                return cell.getStringCellValue().trim();
            case BLANK:
                return "";
            case ERROR:
                return FormulaError.forInt(cell.getErrorCellValue()).getString();
            case BOOLEAN:
                return cell.getBooleanCellValue() ? "TRUE" : "FALSE";
            case NUMERIC:
                if (HSSFDateUtil.isCellDateFormatted(cell)) {
                    Date date = cell.getDateCellValue();
                    if (dateFormat.format(date).startsWith("1899-12-31"))
                        return dateFormat.format(date).split(" ")[1];
                    else
                        return dateFormat.format(date);
                }
                //此处可根据需要,构造DecimalFormat对象、
                return formatter.formatCellValue(cell, evaluator);
            default:
                return "";
        }
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
                return cell.getStringCellValue().trim();
            case BLANK:
                return null;
            case ERROR:
                return cell.getErrorCellValue();
            case FORMULA:
                return cell.getStringCellValue().trim();
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
     *  设置Cell的值、
     * @param cell cell
     * @param value value
     */
    public void setCellValue(Cell cell, Object value) {
        if (value instanceof String) {
            cell.setCellValue(value.toString());
            return;
        }
        if (value instanceof Date) {
            //需要格式化、
            cell.setCellValue((Date) value);
            return;
        }
        if (value instanceof Calendar) {
            cell.setCellValue((Calendar) value);
            return;
        }
        if (value instanceof Boolean) {
            cell.setCellValue((Boolean) value);
            return;
        }
        if (value instanceof Double) {
            cell.setCellValue((Double) value);
            return;
        }
        if (value instanceof RichTextString) {
            cell.setCellValue((RichTextString) value);
            return;
        }
        cell.setCellValue(value.toString());
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
     * 解析sheet的内容,不包含图片、(没有处理表格标题所占的Row)
     * 超链接使用 value = name:url 的格式存在Map中、
     *
     * @param sheet     工作簿、xls,xlsx均可解析、
     * @param evaluator 公式计算器. 遇到单元格是FORMULA类型,传入此值,获得计算后的值,如果为null、则表示获取公式的表达式、
     * @return mapList Excel表格的每行数据,不包含表头
     */
    public List<Map<String, Object>> parseSimpleExcel(Sheet sheet, @Nullable FormulaEvaluator evaluator) {
        List<Map<String, Object>> dataList = new ArrayList<>();
        int firstRowNum = sheet.getFirstRowNum();
        Row row0 = sheet.getRow(firstRowNum);
        if (row0 == null)//空白表格：row0 == null
            return dataList;
        for (Row row : sheet) {
            //不获取表头的数据、(没处理表格标题所占Row,默认第一行是表头)
            if (row.getRowNum() == firstRowNum)
                continue;
            //判断是否是空行、
            if (isRowEmpty(row))
                continue;
            Map<String, Object> map = new HashMap<>();
            for (Cell cell : row) {
                //获取单元格格式化后的值
                String formatValue = getStringCellValue(cell, evaluator);
                //获取当前单元格的坐标
                CellAddress address = cell.getAddress();
                if (cell.getHyperlink() != null) {
                    String linkAddr = cell.getHyperlink().getAddress();
                    map.put(row0.getCell(address.getColumn()).getStringCellValue(), formatValue + ":" + linkAddr);//key是表头、
                    continue;
                }
                //cell.getCellComment();//获取标注、注释
                map.put(row0.getCell(address.getColumn()).getStringCellValue(), formatValue);//key是表头、
            }
            dataList.add(map);
        }
        return dataList;
    }

    /**
     * 解析sheet的内容,包含图片、(没有处理表格标题所占的Row)
     * 超链接使用 value = name:url 的格式存在Map中、
     *
     * @param sheet     xls格式工作簿
     * @param evaluator 公式计算器、可以为null、
     * @return mapList Excel表格的每行数据,不包含表头行
     */
    public List<Map<String, Object>> parseComplexXls(HSSFSheet sheet, @Nullable FormulaEvaluator evaluator) {
        List<Map<String, Object>> dataList = new ArrayList<>();
        Map<CellAddress, PictureData> pictureDataMap = getHSSFPictureData(sheet);
        int firstRowNum = sheet.getFirstRowNum();
        Row row0 = sheet.getRow(firstRowNum);
        if (row0 == null)//空白表格：row0 == null
            return dataList;
        for (Row row : sheet) {
            //不获取表头的数据、(没处理表格标题所占Row,默认第一行是表头)
            if (row.getRowNum() == firstRowNum)
                continue;
            //判断是否是空行、
            if (isRowEmpty(row))
                continue;
            Map<String, Object> map = new HashMap<>();
            for (Cell cell : row) {
                //获取单元格格式化后的值
                String formatValue = getStringCellValue(cell, evaluator);
                //获取当前单元格的坐标
                CellAddress address = cell.getAddress();
                if (cell.getHyperlink() != null) {
                    String linkAddr = cell.getHyperlink().getAddress();
                    map.put(row0.getCell(address.getColumn()).getStringCellValue(), formatValue + ":" + linkAddr);//key是表头、
                    continue;
                }
                //cell.getCellComment();//获取标注、注释
                //cell.getCellFormula();//获取单元格的公式
                map.put(row0.getCell(address.getColumn()).getStringCellValue(), formatValue);//key是表头、
            }
            if (pictureDataMap != null && !pictureDataMap.isEmpty()) {
                pictureDataMap.keySet().stream().filter(address -> address.getRow() == row.getRowNum())
                        .forEach(address -> map.put(row0.getCell(address.getColumn()).getStringCellValue(), pictureDataMap.get(address)));
            }
            dataList.add(map);
        }
        return dataList;
    }

    /**
     * 获取当前sheet中的每个Cell内的图片数据、(图片所占范围超过一个Cell、则不获取)
     *
     * @param sheet xls格式的Excel工作簿、
     * @return 当前sheet中的图片数据、
     */
    private Map<CellAddress, PictureData> getHSSFPictureData(HSSFSheet sheet) {
//        List<HSSFPictureData> hssfPictureDataList = book.getAllPictures(); //无法定位Picture、
        HSSFPatriarch patriarch = sheet.createDrawingPatriarch();
        Map<CellAddress, PictureData> pictureDataMap = new HashMap<>();
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
                /*
                 *   0<= dx <=1023
                 *   0<= dy <=255
                 *   锚点的起点：左上角(dx1,dy1)、
                 *   锚点的终点：右上角(dx2,dy2)、
                 * */
                //只取图片在一个单元格内的、
                if ((col1 == col2 && row1 == row2) || (row1 + 1 == row2 && col1 + 1 == col2 && dx2 == 0 && dy2 == 0)) {
                    CellAddress cellAddress = new CellAddress(row1, col1);
                    pictureDataMap.put(cellAddress, picture.getPictureData());
                }
            }
        }
        return pictureDataMap;
    }

    /**
     * 异步保存Excel中读取数据的图片、
     * 如果需要获取存储路径,返回值改为Future<String>对象即可、
     *
     * @param absPath
     * @param data
     */
    @Async
    public void writeImage(@NonNull String absPath, @NonNull PictureData data) {
        if (StringUtils.isEmpty(absPath) || data == null)
            throw new RuntimeException("absPath 和 data 不能为null.");
        BufferedImage bufferedImage;
        ByteArrayInputStream inputStream = null;
        FileOutputStream fileOutputStream = null;
        BufferedOutputStream bufferedOutputStream = null;
        try {
            inputStream = new ByteArrayInputStream(data.getData());
            bufferedImage = ImageIO.read(inputStream);
            fileOutputStream = new FileOutputStream(absPath + UUID.randomUUID().toString() + "." + data.getMimeType().split("/")[1]);
            bufferedOutputStream = new BufferedOutputStream(fileOutputStream);
            ImageIO.write(bufferedImage, data.getMimeType().split("/")[1], bufferedOutputStream);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            closeStream(inputStream, fileOutputStream, bufferedOutputStream);
        }
    }

    /**
     * 判断此行是否为空行,有空格
     *
     * @param row
     * @return true OR false
     */
    private boolean isRowEmpty(Row row) {
        for (Cell cell : row) {
            if (cell != null && cell.getCellTypeEnum() != CellType.BLANK && !StringUtils.isEmpty(getStringCellValue(cell)))
                return false;
        }
        return true;
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

    public void closeStream(InputStream inputStream, OutputStream... outputStream) {
        if (outputStream != null)
            Arrays.asList(outputStream).forEach(ops -> {
                try {
                    ops.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            });
        if (inputStream != null) {
            try {
                inputStream.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    public void closeStream(OutputStream outputStream, InputStream... inputStream) {
        if (inputStream != null)
            Arrays.asList(inputStream).forEach(ips -> {
                try {
                    ips.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            });
        if (outputStream != null) {
            try {
                outputStream.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }
}
