package com.lxf.poi.util;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.stereotype.Component;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

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
}
