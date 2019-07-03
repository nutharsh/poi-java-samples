package com.adarsh.excel;

import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.File;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

@Slf4j
public class WorkbookRowReader {
  public static Map<Integer, String> getMap(String path, String sheetName, int upcCellNum) throws Exception {
    Map<Integer, String> maps = new HashMap<>();
    final org.apache.poi.ss.usermodel.Workbook wb = WorkbookFactory.create(new File(path));
    final Sheet sheet0 = wb.getSheet(sheetName);
    final Iterator<Row> iterator = sheet0.iterator();
    while ( iterator.hasNext() ) {
      final Row row = iterator.next();
      final Cell cell = row.getCell(upcCellNum, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
      final CellType cellTypeEnum = cell.getCellTypeEnum();
      String stringCellValue = "";
      if( cellTypeEnum == CellType.STRING ) {
        stringCellValue = cell.getStringCellValue();
      } else if( cellTypeEnum == CellType.NUMERIC ) {
        final String formatted = String.format("%f", cell.getNumericCellValue());
        stringCellValue = formatted.substring(0, formatted.indexOf('.'));
      }
      if( !StringUtils.isBlank(stringCellValue) ) {
        String org = stringCellValue;
        stringCellValue = stringCellValue.replaceAll(" ", "");
        stringCellValue = stringCellValue.replaceAll("-", "");
        stringCellValue = stringCellValue.replaceAll("'", "");
        stringCellValue = stringCellValue.replaceAll("<", "");
        stringCellValue = stringCellValue.replaceAll(">", "");
        stringCellValue = stringCellValue.replaceAll("\\r\\n|\\r|\\n", "");
        stringCellValue = stringCellValue.replace("\u00A0", "");
        maps.put(row.getRowNum(), stringCellValue.trim());
        if( org.length() != stringCellValue.length() ) {
          System.out.println(org + " != " + stringCellValue);
        }
      }
      log.info("row={}", row.getRowNum());
    }
    return maps;
  }
}
