package com.lxf.poi.util;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.stereotype.Component;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * @author 小66
 * @create 2019-08-12 21:42
 **/
@Component
@Slf4j
public class POIExcelUtil {

    public void closeStream(Workbook workbook) throws IOException {
        if (workbook != null)
            workbook.close();
    }

    public void closeStream(Workbook workbook, FileOutputStream outputStream) throws IOException {
        if (outputStream != null)
            outputStream.close();
        if (workbook != null)
            workbook.close();
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
