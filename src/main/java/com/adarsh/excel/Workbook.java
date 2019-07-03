package com.adarsh.excel;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFPicture;
import org.apache.poi.hssf.usermodel.HSSFShape;
import org.apache.poi.hssf.usermodel.HSSFShapeGroup;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFShapeGroup;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

/**
 * @author Adarsh Thimmappa
 */
@Slf4j
public class Workbook {
  public static void main(String[] args) throws Exception {
    if( args.length != 5 ) {
      System.err.println("usage: ...Workbook <file-absolute-path> <upc-cell-num> <dept-name> <file_name>");
      System.exit(-1);
    }
    String path = args[0];
    String sheetNbr = args[1];
    Integer upcCellNum = Integer.parseInt(args[2]);
    String folder1 = args[3];
    String folder2 = args[4];
    final String parent = "pics/" + folder1 + "/" + folder2;
    File filePrefix = new File(parent);
    filePrefix.mkdirs();
    final Map<Integer, String> maps = WorkbookRowReader.getMap(path, sheetNbr, upcCellNum);
    final org.apache.poi.ss.usermodel.Workbook wb = WorkbookFactory.create(new File(path));
    final Sheet currentSheet = wb.getSheet(sheetNbr);
    if( currentSheet instanceof HSSFSheet ) {
      processHssfSheet(parent, maps, (HSSFSheet) currentSheet);
    } else if( currentSheet instanceof XSSFSheet ) {
      processXssfSheet(parent, maps, (XSSFSheet) currentSheet);
    }
  }

  private static void processXssfSheet(String parent,
                                       Map<Integer, String> maps,
                                       XSSFSheet currentSheet) throws IOException {
    XSSFSheet xsheet = currentSheet;
    List<XSSFShape> xssfShapes = xsheet.createDrawingPatriarch().getShapes();
    for ( XSSFShape xssfShape : xssfShapes ) {
      if( !(xssfShape instanceof XSSFPicture) ) {
        continue;
      }
      if( xssfShape instanceof XSSFPicture ) {
        XSSFPicture inpPic = (XSSFPicture) xssfShape;
        XSSFClientAnchor anchor = inpPic.getClientAnchor();
        final String mimeType = inpPic.getPictureData().getMimeType();
        String name = maps.get(anchor.getRow1()) + "." + mimeType.split("/")[1];
        File file = new File(parent, name);
        final FileOutputStream fileOutputStream = new FileOutputStream(file);
        fileOutputStream.write(inpPic.getPictureData().getData());
        fileOutputStream.flush();
        fileOutputStream.close();
      } else if( xssfShape instanceof XSSFShapeGroup ) {
        XSSFShapeGroup xssfShapeGroup = (XSSFShapeGroup) xssfShape;
        final XSSFClientAnchor anchor = (XSSFClientAnchor) xssfShapeGroup.getAnchor();
        final Iterator<XSSFShape> iterator = xssfShapeGroup.iterator();
        int x = 0;
        while ( iterator.hasNext() ) {
          String name = maps.get(anchor.getRow1()) + "-" + x;
          XSSFShape shape = iterator.next();
          XSSFPicture pic = (XSSFPicture) shape;
          final String mimeType = pic.getPictureData().getMimeType();
          name = name + "." + mimeType.split("/")[1];
          File file = new File(parent, name);
          final FileOutputStream fileOutputStream = new FileOutputStream(file);
          fileOutputStream.write(pic.getPictureData().getData());
          fileOutputStream.flush();
          fileOutputStream.close();
          x++;
        }
      } else {
        log.info("***unable to recognise {}***", xssfShape.getClass().getName());
      }
    }
  }

  private static void processHssfSheet(String parent,
                                       Map<Integer, String> maps,
                                       HSSFSheet currentSheet) throws IOException {
    HSSFSheet hsheet = currentSheet;
    final List<HSSFShape> hssfShapes = hsheet.createDrawingPatriarch().getChildren();
    for ( HSSFShape hssfShape : hssfShapes ) {
      if( hssfShapes.isEmpty() ) {
        continue;
      }
      if( !(hssfShape instanceof HSSFPicture) ) {
        continue;
      }
      if( hssfShape instanceof HSSFPicture ) {
        HSSFPicture inpPic = (HSSFPicture) hssfShape;
        HSSFClientAnchor anchor = inpPic.getClientAnchor();
        final String mimeType = inpPic.getPictureData().getMimeType();
        if( maps.get(anchor.getRow1()) == null ) {
          continue;
        }
        String name = maps.get(anchor.getRow1()) + "-" + anchor.getRow1() + "." + mimeType.split("/")[1];
        File file = new File(parent, name);
        final FileOutputStream fileOutputStream = new FileOutputStream(file);
        fileOutputStream.write(inpPic.getPictureData().getData());
        fileOutputStream.flush();
        fileOutputStream.close();
      } else if( hssfShape instanceof HSSFShapeGroup ) {
        HSSFShapeGroup hssfShapeGroup = (HSSFShapeGroup) hssfShape;
        final HSSFClientAnchor anchor = (HSSFClientAnchor) hssfShapeGroup.getAnchor();
        final String arg = anchor.getRow1() + "-" + anchor.getCol1();
        final Iterator<HSSFShape> iterator = hssfShapeGroup.iterator();
        int x = 0;
        while ( iterator.hasNext() ) {
          if( maps.get(anchor.getRow1()) == null ) {
            continue;
          }
          String name = maps.get(anchor.getRow1()) + "-" + anchor.getRow1() + "-" + x;
          HSSFShape shape = iterator.next();
          HSSFPicture pic = (HSSFPicture) shape;
          final String mimeType = pic.getPictureData().getMimeType();
          name = name + "." + mimeType.split("/")[1];
          File file = new File(parent, name);
          final FileOutputStream fileOutputStream = new FileOutputStream(file);
          fileOutputStream.write(pic.getPictureData().getData());
          fileOutputStream.flush();
          fileOutputStream.close();
          x++;
        }
      }
    }
  }
}
