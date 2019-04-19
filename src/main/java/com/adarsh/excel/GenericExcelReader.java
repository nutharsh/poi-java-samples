package com.adarsh.excel;

import com.adarsh.excel.exception.ExcelReadException;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.InputStream;

/**
 * Generic Excel Reader API
 *
 * @author Adarsh Thimmappa
 */
@Slf4j
public class GenericExcelReader {

  /**
   * API to read file input stream
   *
   * @param inputStream - input stream instance
   * @return - poi workbook instance
   * @throws ExcelReadException - occurs when excel read fails for any reason
   */
  public static org.apache.poi.ss.usermodel.Workbook getWorkbook(InputStream inputStream) throws ExcelReadException {
    try {
      return WorkbookFactory.create(inputStream);
    } catch (Exception ex) {
      log.error("unable to read excel contents.", ex);
      throw new ExcelReadException(ex);
    }
  }
}
