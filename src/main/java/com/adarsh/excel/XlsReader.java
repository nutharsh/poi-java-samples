package com.adarsh.excel;

import com.adarsh.excel.exception.ExcelReadException;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.InputStream;

/**
 * API to read xls format
 *
 * @author Adarsh Thimmappa
 */
@Slf4j
public class XlsReader {

  /**
   * API accepts xls content as input stream and return Workbook instance
   *
   * @param inputStream - inputStream instance
   * @return - xls specific workbook instance - HSSFWorkbook
   * @throws ExcelReadException - thrown for any sort of read excel failure
   */
  public static HSSFWorkbook readXlsWorkbook(InputStream inputStream) throws
          ExcelReadException {
    try {
      return new HSSFWorkbook(inputStream);
    } catch (Exception ex) {
      log.error("unable to read xls contents.", ex);
      throw new ExcelReadException(ex);
    }
  }
}
