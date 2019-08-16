package com.lxf.poi.util;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellUtil;
import org.springframework.lang.NonNull;
import org.springframework.lang.Nullable;
import org.springframework.scheduling.annotation.Async;
import org.springframework.stereotype.Component;
import org.springframework.util.MimeTypeUtils;
import org.springframework.util.StringUtils;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.*;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * 操作Excel的工具类、没有处理具有表格标题的Excel文件、 getTableTitleRegion()可以获取表格标题的范围、
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
     * 设置Cell的值、
     * Date类型："yyyy-mm-dd HH:mm:ss"、
     * 没有处理Calendar类型的数据、
     * LocalDate类型:Excel不支持、会被转为String、
     *
     * @param cell  cell
     * @param value value
     */
    public void setCellValue(Cell cell, Object value) {
        SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-mm-dd HH:mm:ss");
        if (value instanceof String) {
            cell.setCellValue(value.toString());
            return;
        }
        if (value instanceof Date) {
            //需要格式化、
            String format;
            if (dateFormat.format(value).startsWith("1899-12-31"))
                format = "h:mm:ss";
            else
                format = "yyyy-mm-dd h:mm:ss";
            CellStyle style = getStyle(cell.getRow().getSheet().getWorkbook(), format);
            cell.setCellStyle(style);
            cell.setCellValue((Date) value);
            return;
        }
        if (value instanceof Calendar) {
            //需要处理Calendar类型的格式化、
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
     * @param alignment   水平位置、默认：HorizontalAlignment.CENTER
     * @param borderStyle 边框样式, 默认：BorderStyle.THIN
     * @param borderColor 边框颜色, 默认：IndexedColors.BLACK.index
     * @param wrapText    是否换行, 默认：false
     * @param lock        是否锁定, 默认：false
     * @param format      DataFormat对象的format格式、
     * @return CellStyle
     */
    public CellStyle getStyle(Workbook workbook, Font font, HorizontalAlignment alignment,
                              BorderStyle borderStyle, short borderColor, boolean wrapText, boolean lock, @Nullable String format) {
        CellStyle style = workbook.createCellStyle();

        style.setAlignment(alignment);
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

        if (format!=null)
            style.setDataFormat(workbook.createDataFormat().getFormat(format));

        return style;
    }

    /**
     * 获取自定义的默认样式：水平居中,黑色-细的(THIN)边框,不换行,不锁定,自定义默认字体(Consolas)、
     *
     * @param workbook workbook
     * @return CellStyle
     */
    public CellStyle getStyle(Workbook workbook) {
        return getStyle(workbook, getFont(workbook), HorizontalAlignment.CENTER, BorderStyle.THIN,
                IndexedColors.BLACK.index, false, false,null);
    }

    /**
     * 获取自定义的默认样式：水平居中,黑色-细的(THIN)边框,不换行,不锁定,自定义默认字体(Consolas)、
     *
     * @param workbook workbook
     * @param format   DataFormat对象的format格式、
     * @return CellStyle
     */
    public CellStyle getStyle(Workbook workbook, String format) {
        return getStyle(workbook, getFont(workbook), HorizontalAlignment.CENTER, BorderStyle.THIN,
                IndexedColors.BLACK.index, false, false,format);
    }

    /**
     * 获取默认样式、黑色-细的(THIN)边框,不换行,不锁定、
     * 字体样式:需要自定义、
     *
     * @param workbook  workbook
     * @param font      自定义样式、
     * @param alignment 水平位置
     * @return CellStyle
     */
    public CellStyle getStyle(Workbook workbook, Font font, @Nullable HorizontalAlignment alignment) {
        if (alignment == null) alignment = HorizontalAlignment.CENTER;
        return getStyle(workbook, font, alignment, BorderStyle.THIN,
                IndexedColors.BLACK.index, false, false,null);
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
        if (fontHeightInPoints > 0)
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
     * 获取自定义的默认样式、默认：Consolas、黑色、不加粗、没有下划线、不斜体
     * 注意：不要循环创建字体样式,尽量重用、
     *
     * @param workbook           workbook
     * @param fontHeightInPoints 字号大小、
     * @param bold               是否加粗
     * @return Font
     */
    public Font getFont(Workbook workbook, int fontHeightInPoints, boolean bold) {
        return getFont(workbook, fontHeightInPoints, "Consolas",
                IndexedColors.BLACK.index, bold, false, false, Font.U_NONE);
    }

    /**
     * 获取自定义的默认样式、默认：13字号、Consolas、不斜体
     * 注意：不要循环创建字体样式,尽量重用、
     *
     * @param workbook  workbook
     * @param color     eg:IndexedColors.GREEN.index,HSSFColor.RED.index
     * @param bold      是否加粗
     * @param underLine underLine的样式,默认:Font.U_NONE
     * @return Font
     */
    public Font getFont(Workbook workbook, short color, boolean bold, @Nullable Byte underLine) {
        if (underLine == null) underLine = Font.U_NONE;
        return getFont(workbook, 13, "Consolas",
                color, bold, false, false, underLine);
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
     * 获取当前sheet文件中合并的单元格是不是表格的标题
     *
     * @param sheet sheet
     * @return CellRangeAddress
     */
    public CellRangeAddress getTableTitleRegion(Sheet sheet) {
        List<CellRangeAddress> rangeAddressList = sheet.getMergedRegions();
        if (rangeAddressList == null || rangeAddressList.size() <= 0)
            return null;
        for (CellRangeAddress address : rangeAddressList) {
            int firstRow = address.getFirstRow();
            //int lastRow = address.getLastRow();
            int firstColumn = address.getFirstColumn();
            //int lastColumn = address.getLastColumn();
            if (firstRow == 0 && firstColumn == 0) {//此时这个合并的单元格是表格的标题、(正常情况下)
                return address;
            }
        }
        return null;
    }

    /**
     * 异步创建Excel文件、并补全数据、
     *
     * @param absFilePath       生成的文件路径、
     * @param picturePrefixPath 图片文件夹的存储路径、(没有图片,则为null)
     * @param tableHeader       表头字段名、
     * @param mapList           数据
     * @param splitLink         超链接分隔符、(没有超链接,则为null)
     * @throws Exception IO
     */
    @Async
    public void createXlsAndInsertData(String absFilePath, @Nullable String picturePrefixPath, List<String> tableHeader, List<Map<String, Object>> mapList,
                                       @Nullable Float rowHeightInPoints, @Nullable Integer defaultColumnWidth, @Nullable String splitLink) throws Exception {
        if (StringUtils.isEmpty(splitLink)) splitLink = "---";

        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheetAt = workbook.createSheet();
        //sheetAt.setDefaultRowHeightInPoints(36);//对当前新增的行、没有效果.
        if (defaultColumnWidth != null) sheetAt.setDefaultColumnWidth(defaultColumnWidth);//设置列宽、

        CellStyle headerStyle = getStyle(workbook, getFont(workbook, 16, true), HorizontalAlignment.CENTER);
        CellStyle bodyStyle = getStyle(workbook);
        Font linkFont = getFont(workbook, IndexedColors.BLUE.index, false, Font.U_SINGLE);
        CellStyle linkStyle = getStyle(workbook, linkFont, HorizontalAlignment.LEFT);

        HSSFRow row0 = sheetAt.createRow(0);
        if (rowHeightInPoints != null) row0.setHeightInPoints(rowHeightInPoints + 6);//设置表头行高、
        for (int i = 0; i < tableHeader.size(); i++) {
            CellUtil.createCell(row0, i, tableHeader.get(i), headerStyle);
            //sheetAt.autoSizeColumn(i);自动调整宽度,效果不好、可以在Style中设置换行、
        }
        createRowsWithData(picturePrefixPath, mapList, splitLink, sheetAt, bodyStyle, linkStyle, rowHeightInPoints);

        workbook.write(new File(absFilePath));
        workbook.close();
    }

    /**
     * 不可异步、否则插入数据之前,流已被关闭、
     * <p>
     * 使用默认的bodyStyle和linkStyle参数、
     *
     * @param picturePrefixPath 图片存储路径 (正常是多张图片是统一前缀的)
     * @param mapList           参数
     * @param splitLink         超链接拼接符、
     * @param sheetAt           sheet
     * @param rowHeightInPoints 行高、如果不传,则获取最后一行的行高、
     * @throws IOException IO
     */
    private void createRowsWithData(String picturePrefixPath, List<Map<String, Object>> mapList, String splitLink, HSSFSheet sheetAt, @Nullable Float rowHeightInPoints) throws IOException {
        HSSFWorkbook workbook = sheetAt.getWorkbook();
        CellStyle bodyStyle = getStyle(workbook);
        Font linkFont = getFont(workbook, IndexedColors.BLUE.index, false, Font.U_SINGLE);
        CellStyle linkStyle = getStyle(workbook, linkFont, HorizontalAlignment.LEFT);
        createRowsWithData(picturePrefixPath, mapList, splitLink, sheetAt, bodyStyle, linkStyle, rowHeightInPoints);
    }

    /**
     * 不可异步、否则插入数据之前,流已被关闭、
     *
     * @param picturePrefixPath 图片路径前缀
     * @param mapList           数据
     * @param splitLink         分隔符
     * @param sheetAt           工作簿
     * @param bodyStyle         单元格样式、
     * @param linkStyle         超链接样式、
     * @param rowHeightInPoints 行高、如果为value = null或 value = <= 0,则获取最后一行的行高、
     * @throws IOException IO
     */
    private void createRowsWithData(String picturePrefixPath, List<Map<String, Object>> mapList, String splitLink,
                                    HSSFSheet sheetAt, CellStyle bodyStyle, CellStyle linkStyle, Float rowHeightInPoints) throws IOException {
        HSSFWorkbook workbook = sheetAt.getWorkbook();
        HSSFPatriarch patriarch = sheetAt.createDrawingPatriarch();
        HSSFRow row0 = sheetAt.getRow(sheetAt.getFirstRowNum());
        HSSFRow lastRow = sheetAt.getRow(sheetAt.getLastRowNum());
        if (lastRow == null) throw new RuntimeException(sheetAt.getSheetName() + " is empty");
        if (rowHeightInPoints == null || rowHeightInPoints <= 0) rowHeightInPoints = lastRow.getHeightInPoints();
        HSSFCellStyle lastRowStyle = lastRow.getRowStyle();

        int lastRowNum = sheetAt.getLastRowNum();
        lastRowNum++;
        for (int i = 0; i < mapList.size(); i++) {
            HSSFRow row = sheetAt.createRow(i + lastRowNum);
            row.setHeightInPoints(rowHeightInPoints);
            //复制最后一行(原始)的行属性、
            if (!isRowEmpty(lastRow) && lastRowStyle != null) {
                row.setRowStyle(lastRowStyle);
            }
            Map<String, Object> map = mapList.get(i);
            for (int column = 0; column < map.entrySet().size(); column++) {
                //表头、
                String key = row0.getCell(column).getStringCellValue();
                Object value = map.get(key);
                if (value == null || value.equals("")) continue;

                if (patriarch != null && value.toString().startsWith(picturePrefixPath)) {//图片路径、
                    String absPicturePath = value.toString();
                    /*
                     * TODO: 此处使用异步,会导致图片没有写入的时候,流已被关闭(其他文字数据正常写入)、
                     *       后续使用NIO试试、
                     * */
                    createPicture(workbook, patriarch, row, column, absPicturePath);

                } else if (value.toString().contains(splitLink)) {//超链接
                    String[] strings = value.toString().split(splitLink);//超链接需要指定固定的格式、
                    HSSFCell cell = row.createCell(column);
                    cell.setCellStyle(linkStyle);
                    HSSFHyperlink hyperlink = workbook.getCreationHelper().createHyperlink(HyperlinkType.URL);
                    hyperlink.setAddress(strings[1]);
                    //setShortFilename()没生效、如果不进行cell.setCellValue(strings[0]);则显示空白
                    hyperlink.setShortFilename(strings[0]);
                    cell.setHyperlink(hyperlink);
                    cell.setCellValue(strings[0]);//显示超链接的名称、
                } else {
                    HSSFCell cell = row.createCell(column);
                    if (isRowEmpty(lastRow) || lastRowStyle == null)
                        cell.setCellStyle(bodyStyle);
                    setCellValue(cell, value);
                }
            }
        }
    }

    /**
     * 异步插入数据、
     * <p>
     * 在第一个不为空(有数据)的sheet工作簿的最后插入数据、
     *
     * @param absFilePath       Excel文件路径
     * @param picturePrefixPath 图片存储路径的前缀、(类似根路径)
     * @param mapList           注意：如果存在超链接、使用"---" 拼接, 即value = cellValue---address 的格式存在Map中、
     * @param splitLink         超链接的拼接字符串、默认"---"
     * @throws IOException
     */
    @Async
    public void addNonEmptyRow(String absFilePath, String picturePrefixPath, List<Map<String, Object>> mapList, @Nullable String splitLink) throws IOException {
        if (StringUtils.isEmpty(splitLink)) splitLink = "---";

        FileInputStream fileInputStream = new FileInputStream(absFilePath);
        HSSFWorkbook workbook = new HSSFWorkbook(fileInputStream);
        boolean flag = true;
        boolean hasNext = workbook.iterator().hasNext();
        //TODO: 此处如果使用for(T item : expr)方法循环、使用return;方式跳出循环,数据不会写入到文件中、
        while (hasNext && flag) {
            Sheet sheet = workbook.iterator().next();
            if (!(sheet instanceof HSSFSheet)) continue;
            HSSFSheet sheetAt = (HSSFSheet) sheet;
            HSSFRow row0 = sheetAt.getRow(sheetAt.getFirstRowNum());
            if (row0 == null) continue;//此sheet工作簿为空(没有数据)、
            createRowsWithData(picturePrefixPath, mapList, splitLink, sheetAt, null);
            flag = false;
        }
        workbook.write(new File(absFilePath));
        closeStream(workbook, fileInputStream);
    }

    private void createPicture(HSSFWorkbook workbook, HSSFPatriarch patriarch, HSSFRow row, int column, String absFilePath) throws IOException {
        ByteArrayOutputStream byteArrayOS = readImage(absFilePath);
        //        //获取MimeType、
        String contentType = Files.probeContentType(Paths.get(absFilePath));

        int index;
        switch (contentType) {//TODO: 根据需求添加图片类型的判断、
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
            HSSFClientAnchor anchor = getAnchor(16, 48, 975, 239, column, row.getRowNum(), column, row.getRowNum());
            patriarch.createPicture(anchor, index);
        }
    }

    /**
     * 解析sheet的内容,不包含图片、(没有处理表格标题所占的Row)
     * 超链接使用 value = name---url 的格式存在Map中、
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
                    map.put(row0.getCell(address.getColumn()).getStringCellValue(), formatValue + "---" + linkAddr);//key是表头、
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
     * 超链接使用 "---" 拼接, 即 value = cellValue---address 的格式存在Map中、
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
                    //超链接使用 "---" 拼接 cellValue和address
                    map.put(row0.getCell(address.getColumn()).getStringCellValue(), formatValue + "---" + linkAddr);//key是表头、
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
     * 异步保存Excel中的图片数据到指定的文件夹、
     * 如果需要获取存储路径,返回值改为Future<String>对象即可、
     *
     * @param absPath 文件路径
     * @param data    PictureData对象、通过HSSFPicture获取、
     */
    @Async
    public void writeImage(@NonNull String absPath, @NonNull PictureData data) {
        if (StringUtils.isEmpty(absPath) || data == null)
            throw new RuntimeException("absPath 和 data 不能为null.");
        ByteArrayInputStream inputStream = null;
        FileOutputStream fileOutputStream = null;
        BufferedImage bufferedImage;
        BufferedOutputStream bufferedOutputStream;
        try {
            inputStream = new ByteArrayInputStream(data.getData());
            bufferedImage = ImageIO.read(inputStream);
            fileOutputStream = new FileOutputStream(absPath + UUID.randomUUID().toString() + "." + data.getMimeType().split("/")[1]);
            bufferedOutputStream = new BufferedOutputStream(fileOutputStream);
            //内部会关闭outPutStream、
            ImageIO.write(bufferedImage, data.getMimeType().split("/")[1], bufferedOutputStream);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            closeStream(inputStream, fileOutputStream);
        }
    }

    /**
     * 读取Image,转换为ByteArrayOutputStream、  -->关闭IO流、无法获得展示图片
     *
     * @param filePath 图片绝对路径
     * @return ByteArrayOutputStream
     */
    public ByteArrayOutputStream readImage(String filePath) throws IOException {
        FileInputStream inputStream = null;
        try {
            inputStream = new FileInputStream(filePath);
            BufferedImage bufferedImage = ImageIO.read(inputStream);
            ByteArrayOutputStream byteArrayOS = new ByteArrayOutputStream();
            String mimeType = Files.probeContentType(Paths.get(filePath));
            ImageIO.write(bufferedImage, mimeType.split("/")[1], byteArrayOS);
            bufferedImage.flush();
            return byteArrayOS;
        } finally {
            //byteArrayOS在ImageIo.write()内部已关闭、
            if (inputStream != null) inputStream.close();
        }
    }

    /**
     * 获取画板锚点、HSSFClientAnchor,默认填满单元格、
     * col1 = col2 && row1 =row2 表示占据一个的单元格、
     *
     * @param col1
     * @param row1
     * @param col2
     * @param row2
     * @return HSSFClientAnchor
     */
    public HSSFClientAnchor getAnchor(int col1, int row1, int col2, int row2) {
        return new HSSFClientAnchor(0, 0, 1023, 255, (short) col1, row1, (short) col2, row2);
    }

    /**
     * 获取画板锚点,图片和单元格 有一定的间距、
     * 16, 48, 975, 239
     * dx范围: 0~1023
     * dy范围: 0~255
     *
     * @param dx1
     * @param dy1
     * @param dx2
     * @param dy2
     * @param col1
     * @param row1
     * @param col2
     * @param row2
     * @return
     */
    public HSSFClientAnchor getAnchor(int dx1, int dy1, int dx2, int dy2, int col1, int row1, int col2, int row2) {
        return new HSSFClientAnchor(dx1, dy1, dx2, dy2, (short) col1, row1, (short) col2, row2);
    }

    /**
     * 判断此行是否为空行,有空格
     *
     * @param row 当前row
     * @return true OR false
     */
    public boolean isRowEmpty(Row row) {
        if (row == null) return true;
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
