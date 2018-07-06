/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package take_home_gundu;

import java.awt.Font;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.Date;
import java.util.HashSet;
import java.util.Iterator;
import java.util.Map;
import java.util.TreeMap;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.FontUnderline;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author S530462
 */
public class Take_home_gundu {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) throws FileNotFoundException, IOException {
        // TODO code application logic here
        FileInputStream fis = new FileInputStream("i.xlsx");
        XSSFWorkbook workbook_I = new XSSFWorkbook(fis);
        XSSFSheet sheet_I = workbook_I.getSheetAt(0);

        XSSFWorkbook workbook_O = new XSSFWorkbook();
        XSSFSheet sheet_O = workbook_O.createSheet("output");
        XSSFCellStyle cellstyle = workbook_O.createCellStyle();
        cellstyle.setBorderBottom(BorderStyle.THIN);
        cellstyle.setBorderTop(BorderStyle.THIN);
        cellstyle.setBorderRight(BorderStyle.THIN);
        cellstyle.setBorderLeft(BorderStyle.THIN);
        
        XSSFCellStyle cellstyle1 = workbook_O.createCellStyle();
        
        
        XSSFCellStyle cellstyle2 = workbook_O.createCellStyle();
         cellstyle2.setBorderBottom(BorderStyle.THIN);
        cellstyle2.setBorderTop(BorderStyle.THIN);
        cellstyle2.setBorderRight(BorderStyle.THIN);
        cellstyle2.setBorderLeft(BorderStyle.THIN);
        
        XSSFCellStyle cellstyle3 = workbook_O.createCellStyle();
        cellstyle3.setBorderBottom(BorderStyle.THIN);
        cellstyle3.setBorderTop(BorderStyle.THIN);
        cellstyle3.setBorderRight(BorderStyle.THIN);
        cellstyle3.setBorderLeft(BorderStyle.THIN);
        XSSFCellStyle cellstyle4 = workbook_O.createCellStyle();
        XSSFCellStyle cellstyle5 = workbook_O.createCellStyle();
        
        XSSFColor brown = new XSSFColor(new java.awt.Color(163, 3, 1));
        XSSFColor green = new XSSFColor(new java.awt.Color(173, 204, 152));
        XSSFColor pink = new XSSFColor(new java.awt.Color(244, 204, 182));
        XSSFColor white = new XSSFColor(new java.awt.Color(255, 255, 255));

        XSSFFont font = workbook_O.createFont();
        font.setBold(true);

        XSSFFont font1 = workbook_O.createFont();
        font1.setItalic(true);
        font1.setUnderline(FontUnderline.SINGLE);

        cellstyle.setFont(font);
        cellstyle.setFillForegroundColor(green);

        cellstyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        sheet_O.createRow(1).createCell(1).setCellStyle(cellstyle);
        sheet_O.getRow(1).getCell(1).setCellValue("Name");

        cellstyle1.setFillForegroundColor(green);
        cellstyle1.setFont(font1);
        cellstyle1.setBorderBottom(BorderStyle.THIN);
        cellstyle1.setBorderTop(BorderStyle.THIN);
        cellstyle1.setBorderRight(BorderStyle.THIN);
        cellstyle1.setBorderLeft(BorderStyle.THIN);
        cellstyle1.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        sheet_O.getRow(1).createCell(2).setCellStyle(cellstyle1);
//        sheet_O.getRow((short) 1).createCell((short) 2).setCellStyle(cellstyle1);
        sheet_O.getRow((short) 1).getCell((short) 2).setCellValue(" L, F");
        sheet_O.addMergedRegion(CellRangeAddress.valueOf("C2:D2"));
        
       CellRangeAddress mergedRegions = sheet_O.getMergedRegion(0);
        RegionUtil.setBorderLeft(BorderStyle.THIN, mergedRegions, sheet_O);
        RegionUtil.setBorderRight(BorderStyle.THIN, mergedRegions, sheet_O);
        RegionUtil.setBorderTop(BorderStyle.THIN, mergedRegions, sheet_O);
        RegionUtil.setBorderBottom(BorderStyle.THIN, mergedRegions, sheet_O);
        int rowNum = 4;
        int number = 0;
        DataFormatter formatter = new DataFormatter();

         cellstyle2.setFillForegroundColor(brown);
        cellstyle2.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        XSSFFont font2 = workbook_O.createFont();

        font2.setColor(white);
        font2.setBold(true);

        cellstyle2.setAlignment(HorizontalAlignment.CENTER);
        cellstyle2.setFont(font2);
        
        sheet_O.createRow(3).createCell(0).setCellStyle(cellstyle2);

        for (int x = 1; x < 6; x++) {
            sheet_O.getRow(3).createCell(x).setCellStyle(cellstyle2);
        }
         
         XSSFFont font3 = workbook_O.createFont();
        cellstyle3.setAlignment(HorizontalAlignment.CENTER);

        String sn = formatter.formatCellValue(sheet_I.getRow(0).getCell(0));
        sheet_O.getRow(3).getCell(0).setCellValue(sn);

        String g = formatter.formatCellValue(sheet_I.getRow(0).getCell(2));
        sheet_O.getRow(3).getCell(1).setCellValue(g);

        String cs = formatter.formatCellValue(sheet_I.getRow(0).getCell(5));
        sheet_O.getRow(3).getCell(2).setCellValue(cs);
        System.out.println("c " + cs);

        String a = formatter.formatCellValue(sheet_I.getRow(0).getCell(1));
        sheet_O.getRow(3).getCell(3).setCellValue(a);
        System.out.println("a " + a);

        String ar = formatter.formatCellValue(sheet_I.getRow(0).getCell(3));
        sheet_O.getRow(3).getCell(4).setCellValue(ar);

        String name = formatter.formatCellValue(sheet_I.getRow(0).getCell(4));
        sheet_O.getRow(3).getCell(5).setCellValue(name);

        DateFormat df = new SimpleDateFormat("d-MMM-yy");

        Integer sno, creditscore;
        String album, genre, artist;

        ArrayList<data> dataList = new ArrayList<>();

        for (int i = 1; i < sheet_I.getLastRowNum(); i++) {
            data d = new data();
            Row row = sheet_I.getRow(i);

            for (int j = 1; j <= row.getLastCellNum(); j++) {
                Cell ce = row.getCell(j);
                if (j == 1) {
                    //If you have Header in text It'll throw exception because it won't get NumericValue
                    d.setAlbum_Name(ce.getStringCellValue());
                }
                if (j == 2) {
                    d.setGenre(ce.getStringCellValue());
                }
                if (j == 3) {
                    d.setArtist(ce.getStringCellValue());
                }
                if (j == 4) {
                    d.setRelease_date(df.format(ce.getDateCellValue()));
                }
                if (j == 5) {
                    d.setCredit_score(new Double(ce.getNumericCellValue()).intValue());
                }
            }
            dataList.add(d);
        }   
     
        dataList.sort(Comparator.comparing(data::getGenre).thenComparing(Comparator.comparing(data::getCredit_score).reversed()));  
        String str = dataList.get(0).getGenre();
        boolean flag = true;
       
        cellstyle4.setFillForegroundColor(pink);
        cellstyle4.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cellstyle4.setBorderBottom(BorderStyle.THIN);
        cellstyle4.setBorderTop(BorderStyle.THIN);
        cellstyle4.setBorderRight(BorderStyle.THIN);
        cellstyle4.setBorderLeft(BorderStyle.THIN);
        cellstyle5.setFillForegroundColor(green);
        cellstyle5.setFillPattern(FillPatternType.SOLID_FOREGROUND);
         cellstyle5.setBorderBottom(BorderStyle.THIN);
        cellstyle5.setBorderTop(BorderStyle.THIN);
        cellstyle5.setBorderRight(BorderStyle.THIN);
        cellstyle5.setBorderLeft(BorderStyle.THIN);
        for (data dd : dataList) {
            number++;
            Row row = sheet_O.createRow(rowNum++);
            row.createCell(0).setCellValue(number);
            row.createCell(1).setCellValue(dd.getGenre());
            row.createCell(2).setCellValue(dd.getCredit_score());
            row.createCell(3).setCellValue(dd.getAlbum_Name());
            row.createCell(4).setCellValue(dd.getArtist());
            row.createCell(5).setCellValue(dd.getRelease_date());
            row.getCell(5).setCellStyle(cellstyle3);

            if (!dd.getGenre().equals(str)) {
                str = dd.getGenre();
                if (flag == true) {
                    flag = false;
                } else {
                    flag = true;
                }
            }

            if (flag == true) {
                
                for(int row_a = 0; row_a <6; row_a++ )
                {
                    row.getCell(row_a).setCellStyle(cellstyle5);
//                    row.getCell(1).setCellStyle(cellstyle5);
                       
                }
                
            } else {
                
                for(int row_a = 0; row_a <6; row_a++ )
                {
                    row.getCell(row_a).setCellStyle(cellstyle4);
//                   
                }
            }
        }

       

//        for (int y = 4; y < 28; y++) {
//            sheet_O.getRow(y).getCell(5).setCellStyle(cellstyle3);
//        }
        FileOutputStream out = new FileOutputStream("output.xlsx");
        workbook_O.write(out);
        workbook_O.close();
    }
}
