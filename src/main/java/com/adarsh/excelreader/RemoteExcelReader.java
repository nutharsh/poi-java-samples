package com.adarsh.excelreader;

import lombok.extern.slf4j.Slf4j;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.InputStream;
import java.net.URL;

/**
 * @author Adarsh Thimmappa
 */
@Slf4j
public class RemoteExcelReader {
  public static void main(String[] args) throws Exception {
    String location = "https://teams.wal-mart.com/sites/buycon/Shared%20Documents/pic1555589218855-12-14-15.png";
    final InputStream is = new URL(location).openStream();
    /*final org.apache.poi.ss.usermodel.Workbook workbook = WorkbookFactory.create(is);
    log.info("workbook={}", workbook);*/

    final BufferedImage bufferedImage = ImageIO.read(is);
    log.info("bufferedImage={}", bufferedImage);
  }
}
