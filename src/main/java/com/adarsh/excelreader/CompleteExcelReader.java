package com.adarsh.excelreader;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.Iterator;

/**
 * @author Adarsh Thimmappa
 */
@Slf4j
public class CompleteExcelReader {
  public static void main(String[] args) throws Exception {
    XSSFWorkbook workbook = new XSSFWorkbook(
            "/Users/a0t00gz/buyer-connect/excel-reader/src/main/resources/basic-fields.xlsx");
    final XSSFSheet sheet = workbook.getSheetAt(0);

    //log.info("sheet={}", sheet.getFirstRowNum());

    final Iterator<Row> iterator = sheet.iterator();
    while ( iterator.hasNext() ) {
      final Row row = iterator.next();
      final Iterator<Cell> cellIterator = row.cellIterator();
      while ( cellIterator.hasNext() ) {
        final Cell cell = cellIterator.next();
        log.info("type={} val={} ", cell.getCellTypeEnum(), cell);
        final Hyperlink hyperlink = cell.getHyperlink();
        if( hyperlink != null ) {
          log.info("hyperlink={}", hyperlink.getAddress());
        }
      }
    }

  }
}
