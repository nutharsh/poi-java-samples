package com.adarsh.excel;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFPicture;
import org.apache.poi.hssf.usermodel.HSSFShape;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.io.File;
import java.util.Iterator;
import java.util.List;

/**
 * @author Adarsh Thimmappa
 */
@Slf4j
public class Workbook {
  public static void main(String[] args) throws Exception {
    final org.apache.poi.ss.usermodel.Workbook wb = WorkbookFactory.create(new File(
            "/Users/a0t00gz/buyer-connect/excel-reader/src/main/resources/2019FallBTR-ChileD07copy.wb.xlsx"));

    final Iterator<Sheet> sheetIterator = wb.sheetIterator();
    while ( sheetIterator.hasNext() ) {
      int counter = 0;
      final Sheet currentSheet = sheetIterator.next();
      log.info("sheet={}", currentSheet.getSheetName());
      if( currentSheet instanceof HSSFSheet ) {

        HSSFSheet sheet = (HSSFSheet) currentSheet;
        final List<HSSFShape> hssfShapes = sheet.createDrawingPatriarch().getChildren();
        log.info("hssfShapes={}", hssfShapes.size());
        if( hssfShapes.isEmpty() ) {
          continue;
        }
        for ( HSSFShape hssfShape : hssfShapes ) {
          //log.info("hssf shape={}", hssfShape.getClass());
          if( !(hssfShape instanceof HSSFPicture) ) {
            log.warn("unsupported shape={}", hssfShape);
            continue;
          }
          final HSSFClientAnchor clientAnchor = ((HSSFPicture) hssfShape).getClientAnchor();
          log.info("anchor=r{},c{}", clientAnchor.getRow1(), clientAnchor.getCol1());
          counter++;
        }
      } else if( currentSheet instanceof XSSFSheet ) {

        XSSFSheet sheet = (XSSFSheet) currentSheet;
        List<XSSFShape> xssfShapes = sheet.createDrawingPatriarch().getShapes();
        log.info("xssfShapes={}", xssfShapes.size());
        if( xssfShapes.isEmpty() ) {
          continue;
        }
        for ( XSSFShape xssfShape : xssfShapes ) {
          //log.info("xssf shape={}", xssfShape.getClass());
          if( !(xssfShape instanceof XSSFPicture) ) {
            log.warn("unsupported shape={}", xssfShape.getShapeName());
            continue;
          }
          final XSSFClientAnchor clientAnchor = ((XSSFPicture) xssfShape).getClientAnchor();
          log.info("anchor=r{},c{}", clientAnchor.getRow1(), clientAnchor.getCol1());
          counter++;
        }
      }
      log.info("total pics found={}", counter);
    }


    /*int picIndex = 0;
    for ( PictureData pd : allPictures ) {
      log.info("PictureData={}", pd.suggestFileExtension());

      if( pd instanceof HSSFPictureData ) {
        HSSFPictureData hssfPd = (HSSFPictureData) pd;
        final File pic = new File("pics", "pic" + picIndex++ + "." + pd.suggestFileExtension());
        FileOutputStream fos = new FileOutputStream(pic);
        fos.write(hssfPd.getData());
        fos.flush();
        fos.close();
      }
    }*/
  }
}
