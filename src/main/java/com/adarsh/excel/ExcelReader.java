package com.adarsh.excel;


import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFShapeGroup;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.util.Iterator;
import java.util.List;

/**
 * @author Adarsh Thimmappa
 */
@Slf4j
public class ExcelReader {
  public static void main(String[] args) throws Exception {
    // /Users/a0t00gz/buyer-connect/excel-reader/src/main/resources/2019FallBTR-ChileD07.xlsb
    /*final OPCPackage opcPackage = OPCPackage.open(
            "/Users/a0t00gz/buyer-connect/excel-reader/src/main/resources/2019FallBTR-ChileD07.xlsb");
    XSSFBReader xssfbReader = new XSSFBReader(opcPackage);
    System.out.println(xssfbReader);
    XSSFBReader.SheetIterator it = (XSSFBReader.SheetIterator) xssfbReader.getSheetsData();*/

    XSSFWorkbook workbook = new XSSFWorkbook(
            "/Users/a0t00gz/buyer-connect/excel-reader/src/main/resources/2019FallBTR-ChileD07copy.wb.xlsx");
    final XSSFSheet dataSheet = workbook.getSheetAt(0);
    /*final List<XSSFPictureData> allPictures = workbook.getAllPictures();
    int i = 0;
    String namePrefix = "picture-" + System.currentTimeMillis();
    for ( XSSFPictureData picture : allPictures ) {
      final String mimeType = picture.getMimeType();
      log.info("name={} mimetype={}", picture, mimeType);
      String name = namePrefix + "-" + i;
      //((ZipPackagePart) picture.getPackagePart()).getZipArchive().getName();
      if( mimeType.contains("jpeg") ) {
        name = name + ".jpeg";
      } else if( mimeType.contains("png") ) {
        name = name + ".png";
      } else if( mimeType.contains("jpg") ) {
        name = name + ".jpg";
      }
      File file = new File("pics", name);
      final FileOutputStream fileOutputStream = new FileOutputStream(file);
      fileOutputStream.write(picture.getData());
      fileOutputStream.flush();
      fileOutputStream.close();
      i++;
    }*/

    /*log.info("sheet={}", dataSheet.getSheetName());
    final Iterator<Row> iterator = dataSheet.iterator();
    // skip header row
    iterator.next();
    //iterator.next();
    while ( iterator.hasNext() ) {
      final Row row = iterator.next();
      log.info("last cell={}", row.getLastCellNum());
      final Iterator<Cell> cellIterator = row.cellIterator();
      while ( cellIterator.hasNext() ) {
        final Cell cell = cellIterator.next();
        final XSSFCell xssfCell = (XSSFCell) cell;
        log.info(" index={} cell={} value={}", xssfCell.getColumnIndex(), cell.getCellTypeEnum().name(),
                xssfCell);

      }
    }*/

    XSSFDrawing dp = workbook.getSheetAt(1).createDrawingPatriarch();
    List<XSSFShape> pics = dp.getShapes();
    log.info("total={}", pics.size());
    int p = 0;
    for ( XSSFShape xssfShape : pics ) {
      String picPrefix = "pic" + System.currentTimeMillis();
      if( xssfShape instanceof XSSFPicture ) {
        XSSFPicture inpPic = (XSSFPicture) xssfShape;
        XSSFClientAnchor anchor = inpPic.getClientAnchor();
        final String arg = anchor.getRow1() + "-" + anchor.getCol1();
        /*log.info("col1: " + clientAnchor.getCol1() + ", col2: " + clientAnchor.getCol2() + ", row1: " + clientAnchor
                .getRow1() + ", row2: " + clientAnchor.getRow2());*/
        /*log.info("x1: " + clientAnchor.getDx1() + ", x2: " + clientAnchor.getDx2() + ", y1: " + clientAnchor.getDy1() +
                ", y2: " + clientAnchor.getDy2());*/
        final String mimeType = inpPic.getPictureData().getMimeType();

        String name = picPrefix + "-" + p + "-" + arg + "." + mimeType.split("/")[1];
        File file = new File("pics", name);
        final FileOutputStream fileOutputStream = new FileOutputStream(file);
        fileOutputStream.write(inpPic.getPictureData().getData());
        fileOutputStream.flush();
        fileOutputStream.close();
        p++;
      } else if( xssfShape instanceof XSSFShapeGroup ) {
        log.info("\n\n");
        XSSFShapeGroup xssfShapeGroup = (XSSFShapeGroup) xssfShape;
        final XSSFClientAnchor anchor = (XSSFClientAnchor) xssfShapeGroup.getAnchor();
        final String arg = anchor.getRow1() + "-" + anchor.getCol1();
        log.info("###anchor details={}###", arg);
        final Iterator<XSSFShape> iterator = xssfShapeGroup.iterator();
        int x = 0;
        while ( iterator.hasNext() ) {
          String picName = "grp-pic-" + arg + "-" + x;

          XSSFShape shape = iterator.next();
          XSSFPicture pic = (XSSFPicture) shape;
          final String mimeType = pic.getPictureData().getMimeType();
          picName = picName + "." + mimeType.split("/")[1];
          log.info("pic mime={} length={}", mimeType, pic.getPictureData().getData().length);
          File file = new File("pics", picName);
          final FileOutputStream fileOutputStream = new FileOutputStream(file);
          fileOutputStream.write(pic.getPictureData().getData());
          fileOutputStream.flush();
          fileOutputStream.close();
          x++;
          /*log.info("col1: " + anchor.getCol1() + ", col2: " + anchor.getCol2() + ", row1: " + anchor
                  .getRow1() + ", row2: " + anchor.getRow2());*/
        }
      } else {
        log.info("***unable to recognise {}***", xssfShape.getClass().getName());
      }

    }



    /*inpPic.getShapeName();
    PictureData pict = inpPic.getPictureData();
    FileOutputStream out = new FileOutputStream("pict.jpg");
    byte[] data = pict.getData();
    out.write(data);
    out.close();
    System.out.println("col1: " + clientAnchor.getCol1() + ", col2: " + clientAnchor.getCol2() + ", row1: " + clientAnchor.getRow1() + ", row2: " + clientAnchor.getRow2());
    System.out.println("x1: " + clientAnchor.getDx1() + ", x2: " + clientAnchor.getDx2() +  ", y1: " + clientAnchor.getDy1() +  ", y2: " + clientAnchor.getDy2());*/


  }
}
