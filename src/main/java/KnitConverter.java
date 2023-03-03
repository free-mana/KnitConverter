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
import java.io.*;
import java.util.HashMap;

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
        HashMap<Integer, XSSFColor> colors = new HashMap<>();
        HashMap<Integer, XSSFCellStyle> styles = new HashMap<>();
        Sheet chart = wb.createSheet("chart");
        IndexedColorMap colorMap = new DefaultIndexedColorMap();
        XSSFColor gray = new XSSFColor(new Color(127, 127, 127), colorMap);     // #7F7F7F

        int maxI = 0;
        int maxJ = 0;
        for(int i = 0; i < img.getHeight(); i++){
            maxI = i;
            for(int j = 0; j < img.getWidth(); j++){
                maxJ = j;
                int curColor = img.getRGB(j, i);
                if(!colors.containsKey(curColor)){
                    XSSFColor xssfColor = new XSSFColor(new Color(curColor), colorMap);
                    colors.put(curColor, xssfColor);
                    XSSFCellStyle xssfCellStyle = makeStyle(wb, xssfColor, gray);
                    styles.put(curColor, xssfCellStyle);
                }
            }
        }
        for(int i = 0; i < img.getHeight(); i++){
            maxI = i;
            Row row = chart.createRow(i);
            for(int j = 0; j < img.getWidth(); j++){
                maxJ = j;
                Cell c = row.createCell(j);
                int curColor = img.getRGB(j, i);
                c.setCellStyle(styles.get(curColor));
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
        setBorders(borderColor, result);
        result.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return result;
    }
}
