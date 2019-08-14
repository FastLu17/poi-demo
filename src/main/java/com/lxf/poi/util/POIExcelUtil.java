package com.lxf.poi.util;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.springframework.stereotype.Component;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * 操作Excel的工具类、
 * @author 小66
 * @Description
 * @create 2019-08-12 16:29
 **/
@Component
@Slf4j
public class POIExcelUtil {
    private final String BASE_URL = "C:\\Users\\Administrator\\Desktop\\POI\\";

    public void closeStream(Workbook workbook, FileOutputStream outputStream) throws IOException {
        if (outputStream!=null)
            outputStream.close();
        if (workbook!=null)
            workbook.close();
    }

    public void closeStream(Workbook workbook, FileOutputStream outputStream, FileInputStream inputStream)throws IOException {
        if (outputStream!=null)
            outputStream.close();
        if (workbook!=null)
            workbook.close();
        if (inputStream!=null)
            inputStream.close();
    }

    /**
     * 读入excel的内容转换成字符串
     * @param cell
     * @return
     */
    private  String getStringValueFromCell(Cell cell) {
        SimpleDateFormat sFormat = new SimpleDateFormat("yyyy/MM/dd");
        DecimalFormat decimalFormat = new DecimalFormat("#.#");
        String cellValue = "";
        if(cell == null) {
            return cellValue;
        }
        else if(cell.getCellTypeEnum() == CellType.STRING) {
            cellValue = cell.getStringCellValue();
        }

        else if(cell.getCellTypeEnum() == CellType.NUMERIC) {
            if(HSSFDateUtil.isCellDateFormatted(cell)) {
                double d = cell.getNumericCellValue();
                Date date = HSSFDateUtil.getJavaDate(d);
                cellValue = sFormat.format(date);
            }
            else {
                cellValue = decimalFormat.format((cell.getNumericCellValue()));
            }
        }
        else if(cell.getCellTypeEnum() == CellType.BLANK) {
            cellValue = "";
        }
        else if(cell.getCellTypeEnum() == CellType.BOOLEAN) {
            cellValue = String.valueOf(cell.getBooleanCellValue());
        }
        else if(cell.getCellTypeEnum() == CellType.ERROR) {
            cellValue = "";
        }
        else if(cell.getCellTypeEnum() == CellType.FORMULA) {
            cellValue = cell.getCellFormula().toString();
        }
        return cellValue;
    }
}
