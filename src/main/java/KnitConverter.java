/*
 *  main.java.KnitConverter: a tool for making colorwork charts
 *  Copyright (C) 2023 Alan Freeman
 *
 *  This program is free software: you can redistribute it and/or modify
 *  it under the terms of the GNU General Public License as published by
 *  the Free Software Foundation, either version 3 of the License, or
 *  (at your option) any later version.
 *
 *  This program is distributed in the hope that it will be useful,
 *  but WITHOUT ANY WARRANTY; without even the implied warranty of
 *  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 *  GNU General Public License for more details.
 *
 *  You should have received a copy of the GNU General Public License
 *  along with this program.  If not, see <https://www.gnu.org/licenses/>.
 *
 */

import org.apache.commons.codec.DecoderException;
import org.apache.commons.codec.binary.Hex;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder;
import org.apache.xmlbeans.*;
import org.openxmlformats.schemas.drawingml.x2006.main.CTSRgbColor;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTColors;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTIndexedColors;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTMRUColors;
import org.w3c.dom.Node;
import org.xml.sax.ContentHandler;
import org.xml.sax.SAXException;
import org.xml.sax.ext.LexicalHandler;

import javax.imageio.ImageIO;
import javax.xml.namespace.QName;
import javax.xml.stream.XMLStreamReader;
import java.awt.*;
import java.awt.Color;
import java.awt.image.BufferedImage;
import java.awt.image.Raster;
import java.io.*;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Map;

public class KnitConverter{
    public static void main(String[] args){
        File imgFile = new File("sample.png");
        BufferedImage img;
        try{
            img = ImageIO.read(imgFile);
            getCSV(img);
            makeXLSX(img);
        }catch(IOException e){
            e.printStackTrace();
        }
    }
    public static void getCSV(BufferedImage img, File outFile){
        try{
            PrintWriter pw = new PrintWriter(outFile);
            for(int i = 0; i < img.getHeight(); i++){
                for(int j = 0; j < img.getWidth(); j++){
                    pw.print(img.getRGB(j, i));
                    pw.print(',');
                }
                pw.print('\n');
            }
            pw.close();
        }catch(FileNotFoundException e){
            e.printStackTrace();
        }
    }

    public static void getCSV(BufferedImage img){
        File result = new File("chart.csv");
        getCSV(img, result);
    }

    public static void makeXLSX(BufferedImage img){
        Workbook wb = new XSSFWorkbook();
        Raster raster = img.getData();
        int array[] = new int[3];
        Sheet chart = wb.createSheet("chart");
        XSSFColor gray = new XSSFColor(new Color(128, 128, 128), new DefaultIndexedColorMap());
        XSSFCellStyle white = (XSSFCellStyle)wb.createCellStyle();
        white.setFillForegroundColor(new XSSFColor(new Color(255, 255, 255), new DefaultIndexedColorMap()));
        white.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        white.setBorderBottom(BorderStyle.THIN);
        white.setBorderLeft(BorderStyle.THIN);
        white.setBorderRight(BorderStyle.THIN);
        white.setBorderTop(BorderStyle.THIN);
        white.setBorderColor(XSSFCellBorder.BorderSide.BOTTOM, gray);
        white.setBorderColor(XSSFCellBorder.BorderSide.LEFT, gray);
        white.setBorderColor(XSSFCellBorder.BorderSide.RIGHT, gray);
        white.setBorderColor(XSSFCellBorder.BorderSide.TOP, gray);
        XSSFCellStyle black = (XSSFCellStyle)wb.createCellStyle();
        black.setFillForegroundColor(new XSSFColor(new Color(0,0,0), new DefaultIndexedColorMap()));
        black.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        black.setBorderBottom(BorderStyle.THIN);
        black.setBorderLeft(BorderStyle.THIN);
        black.setBorderRight(BorderStyle.THIN);
        black.setBorderTop(BorderStyle.THIN);
        black.setBorderColor(XSSFCellBorder.BorderSide.BOTTOM, gray);
        black.setBorderColor(XSSFCellBorder.BorderSide.LEFT, gray);
        black.setBorderColor(XSSFCellBorder.BorderSide.RIGHT, gray);
        black.setBorderColor(XSSFCellBorder.BorderSide.TOP, gray);
        int maxI = 0;
        int maxJ = 0;
        for(int i = 0; i < img.getHeight(); i++){
            maxI = i;
            Row r = chart.createRow(i);
            for(int j = 0; j < img.getWidth(); j++){
                maxJ = j;
                Cell c = r.createCell(j);
                array = raster.getPixel(j, i, array);
                String red = Integer.toHexString(array[0]);
                if(red.length() % 2 == 1){
                    red = "0" + red;
                }
                String green = Integer.toHexString(array[1]);
                if(green.length() % 2 == 1){
                    green = "0" + green;
                }
                String blue = Integer.toHexString(array[2]);
                if(blue.length() % 2 == 1){
                    blue = "0" + blue;
                }
                String rgb = red + green + blue;
                //c.setCellValue(rgb);
                if(rgb.equals("ffffff")){
                    c.setCellStyle(white);
                }else{
                    c.setCellStyle(black);
                }
//                try{
//                    byte bytes[] = Hex.decodeHex(rgb);
//                    CTSRgbColor color = CTSRgbColor.Factory.newInstance();
//                    color.setVal(bytes);
//                    System.out.println(Arrays.toString(color.getAlphaArray()));
//                    CTColors ctColors = CTColors.Factory.newInstance();
//                    CTIndexedColors ctIndexedColors = CTIndexedColors.Factory.newInstance();
//                    ctIndexedColors.addNewRgbColor().setRgb(bytes);
//                    ctColors.setIndexedColors(ctIndexedColors);
//                    CustomIndexedColorMap colorMap = CustomIndexedColorMap.fromColors(ctColors);
//                    Map<String, Object> map = new HashMap<>();
//                    map.put(CellUtil.FILL_FOREGROUND_COLOR, new XSSFColor(new java.awt.Color(255, 255, 255), new DefaultIndexedColorMap()));
//                    map.put(CellUtil.FILL_PATTERN, FillPatternType.SOLID_FOREGROUND);
//                    CellUtil.setCellStyleProperties(c, map);
//                }catch(DecoderException e){
//                    e.printStackTrace();
//                }
            }
        }
        int cur = maxI + 1;
        for(int i = 0; i < maxI + 1; i++){
            Row r = chart.getRow(i);
            Cell c = r.createCell(maxJ + 1);
            c.setCellValue(cur);
            cur--;
        }
        cur = maxJ + 1;
        Row lastRow = chart.createRow(maxI + 1);
        for(int j = 0; j < maxJ + 1; j++){
            Cell c = lastRow.createCell(j);
            c.setCellValue(cur);
            cur--;
        }
        for(int i = 0; i < maxJ + 2; i++){
            chart.autoSizeColumn(i);
        }
        try (OutputStream fileOut = new FileOutputStream("workbook.xlsx")) {
            wb.write(fileOut);
        }catch(IOException e){
            e.printStackTrace();
        }
    }
}
