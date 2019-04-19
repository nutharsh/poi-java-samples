package com.adarsh.excel;

import com.adarsh.excel.exception.ExcelReadException;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.InputStream;

/**
 * API to read xlsx format
 *
 * @author Adarsh Thimmappa
 */
@Slf4j
public class XlsxReader {

  /**
   * API accepts xlsx content as input stream and return Workbook instance
   *
   * @param inputStream - inputStream instance
   * @return - xlsx specific workbook instance - XSSFWorkbook
   * @throws ExcelReadException - thrown for any sort of read excel failure
   */
  public static XSSFWorkbook readXlsxWorkbook(InputStream inputStream) throws
          ExcelReadException {
    try {
      return new XSSFWorkbook(inputStream);
    } catch (Exception ex) {
      log.error("unable to read xlsx contents.", ex);
      throw new ExcelReadException(ex);
    }
  }
}
