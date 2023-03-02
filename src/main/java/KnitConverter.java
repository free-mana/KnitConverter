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

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder;

import javax.imageio.ImageIO;
import java.awt.Color;
import java.awt.image.BufferedImage;
import java.awt.image.Raster;
import java.io.*;
import java.util.Arrays;

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
        XSSFWorkbook wb = new XSSFWorkbook();
        Raster raster = img.getData();
        int array[] = new int[3];
        byte cArray[] = new byte[3];
        Sheet chart = wb.createSheet("chart");
        IndexedColorMap colorMap = new DefaultIndexedColorMap();
        XSSFColor white = new XSSFColor(new Color(255, 255, 255), colorMap);    // #FFFFFF
        XSSFColor lightGray = new XSSFColor(new Color(191,191,191), colorMap);  // #DFDFDF
        XSSFColor gray = new XSSFColor(new Color(127, 127, 127), colorMap);     // #7F7F7F
        XSSFColor darkGray = new XSSFColor(new Color(63,63,63), colorMap);      // #3F3F3F
        XSSFColor black = new XSSFColor(new Color(0, 0, 0), colorMap);          // #000000
        XSSFColor red = new XSSFColor(new Color(255, 0, 0), colorMap);          // #FF0000
        XSSFColor scarlet = new XSSFColor(new Color(255,63,0), colorMap);       // #FF3F00
        XSSFColor orange = new XSSFColor(new Color(255, 127, 0), colorMap);     // #FF7F00
        XSSFColor gold = new XSSFColor(new Color(255,191,0), colorMap);         // #FFDF00
        XSSFColor yellow = new XSSFColor(new Color(255, 255, 0), colorMap);     // #FFFF00
        XSSFColor lime = new XSSFColor(new Color(127,255,0), colorMap);         // #7FFF00
        XSSFColor green = new XSSFColor(new Color(0,255,0), colorMap);          // #00FF00
        XSSFColor seaGreen = new XSSFColor(new Color(0,255,127), colorMap);     // #00FF7F
        XSSFColor cyan = new XSSFColor(new Color(0,255,255), colorMap);         // #00FFFF
        XSSFColor cornflower = new XSSFColor(new Color(0,127,255), colorMap);   // #007FFF
        XSSFColor blue = new XSSFColor(new Color(0,0,255), colorMap);           // #0000FF
        XSSFColor indigo = new XSSFColor(new Color(63,0,255), colorMap);        // #3F00FF
        XSSFColor purple = new XSSFColor(new Color(127,0,255), colorMap);       // #7F00FF
        XSSFColor purpleMagenta = new XSSFColor(new Color(191,0,255), colorMap);// #DF00FF
        XSSFColor magenta = new XSSFColor(new Color(255,0,255), colorMap);      // #FF00FF
        XSSFColor magentaRed = new XSSFColor(new Color(255,0,127), colorMap);   // #FF007F
        XSSFCellStyle styleWhite = makeStyle(wb, white, gray);
        XSSFCellStyle styleLightGray = makeStyle(wb, lightGray, darkGray);
        XSSFCellStyle styleGray = makeStyle(wb, gray, black);
        XSSFCellStyle styleDarkGray = makeStyle(wb, darkGray, lightGray);
        XSSFCellStyle styleBlack = makeStyle(wb, black, gray);
        XSSFCellStyle styleRed = makeStyle(wb, red, gray);
        XSSFCellStyle styleScarlet = makeStyle(wb, scarlet, gray);
        XSSFCellStyle styleOrange = makeStyle(wb, orange, gray);
        XSSFCellStyle styleGold = makeStyle(wb, gold, gray);
        XSSFCellStyle styleYellow = makeStyle(wb, yellow, gray);
        XSSFCellStyle styleLime = makeStyle(wb, lime, gray);
        XSSFCellStyle styleGreen = makeStyle(wb, green, gray);
        XSSFCellStyle styleSeaGreen = makeStyle(wb, seaGreen, gray);
        XSSFCellStyle styleCyan = makeStyle(wb, cyan, gray);
        XSSFCellStyle styleCornflower = makeStyle(wb, cornflower, gray);
        XSSFCellStyle styleBlue = makeStyle(wb, blue, gray);
        XSSFCellStyle styleIndigo = makeStyle(wb, indigo, gray);
        XSSFCellStyle stylePurple = makeStyle(wb, purple, gray);
        XSSFCellStyle stylePurpleMagenta = makeStyle(wb, purpleMagenta, gray);
        XSSFCellStyle styleMagenta = makeStyle(wb, magenta, gray);
        XSSFCellStyle styleMagentaRed = makeStyle(wb, magentaRed, gray);

        int maxI = 0;
        int maxJ = 0;
        for(int i = 0; i < img.getHeight(); i++){
            maxI = i;
            Row row = chart.createRow(i);
            for(int j = 0; j < img.getWidth(); j++){
                maxJ = j;
                Cell c = row.createCell(j);
                array = raster.getPixel(j, i, array);

                if(array[0] == -1){
                    cArray[0] = (byte)255;
                }else{
                    cArray[0] = (byte)array[0];
                }
                if(array[1] == -1){
                    cArray[1] = (byte)255;
                }else{
                    cArray[1] = (byte)array[1];
                }
                if(array[2] == -1){
                    cArray[2] = (byte)255;
                }else{
                    cArray[2] = (byte)array[2];
                }
                if(Arrays.equals(cArray, white.getRGB())){
                    c.setCellStyle(styleWhite);
                }else if(Arrays.equals(cArray, lightGray.getRGB())){
                    c.setCellStyle(styleLightGray);
                }else if(Arrays.equals(cArray, gray.getRGB())){
                    c.setCellStyle(styleGray);
                }else if(Arrays.equals(cArray, darkGray.getRGB())){
                    c.setCellStyle(styleDarkGray);
                }else if(Arrays.equals(cArray, black.getRGB())){
                    c.setCellStyle(styleBlack);
                }else if(Arrays.equals(cArray, red.getRGB())){
                    c.setCellStyle(styleRed);
                }else if(Arrays.equals(cArray, scarlet.getRGB())){
                    c.setCellStyle(styleScarlet);
                }else if(Arrays.equals(cArray, orange.getRGB())){
                    c.setCellStyle(styleOrange);
                }else if(Arrays.equals(cArray, gold.getRGB())){
                    c.setCellStyle(styleGold);
                }else if(Arrays.equals(cArray, yellow.getRGB())){
                    c.setCellStyle(styleYellow);
                }else if(Arrays.equals(cArray, lime.getRGB())){
                    c.setCellStyle(styleLime);
                }else if(Arrays.equals(cArray, green.getRGB())){
                    c.setCellStyle(styleGreen);
                }else if(Arrays.equals(cArray, seaGreen.getRGB())){
                    c.setCellStyle(styleSeaGreen);
                }else if(Arrays.equals(cArray, cyan.getRGB())){
                    c.setCellStyle(styleCyan);
                }else if(Arrays.equals(cArray, cornflower.getRGB())){
                    c.setCellStyle(styleCornflower);
                }else if(Arrays.equals(cArray, blue.getRGB())){
                    c.setCellStyle(styleBlue);
                }else if(Arrays.equals(cArray, indigo.getRGB())){
                    c.setCellStyle(styleIndigo);
                }else if(Arrays.equals(cArray, purple.getRGB())){
                    c.setCellStyle(stylePurple);
                }else if(Arrays.equals(cArray, purpleMagenta.getRGB())){
                    c.setCellStyle(stylePurpleMagenta);
                }else if(Arrays.equals(cArray, magenta.getRGB())){
                    c.setCellStyle(styleMagenta);
                }else if(Arrays.equals(cArray, magentaRed.getRGB())){
                    c.setCellStyle(styleMagentaRed);
                }else{
                    System.out.println("unknown color, cArray = " + Arrays.toString(cArray));
                }
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
    private static void setBorders(XSSFColor borderColor, XSSFCellStyle style){
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderColor(XSSFCellBorder.BorderSide.BOTTOM, borderColor);
        style.setBorderColor(XSSFCellBorder.BorderSide.LEFT, borderColor);
        style.setBorderColor(XSSFCellBorder.BorderSide.RIGHT, borderColor);
        style.setBorderColor(XSSFCellBorder.BorderSide.TOP, borderColor);
    }
    private static XSSFCellStyle makeStyle(XSSFWorkbook wb, XSSFColor color, XSSFColor borderColor){
        XSSFCellStyle result = wb.createCellStyle();
        result.setFillForegroundColor(color);
        if(borderColor != null){
            setBorders(borderColor, result);
        }
        result.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return result;
    }
}
