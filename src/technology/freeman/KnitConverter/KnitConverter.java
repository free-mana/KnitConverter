/*
 *  KnitConverter: a tool for making colorwork charts
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

package technology.freeman.KnitConverter;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.PrintWriter;


public class KnitConverter{
    public static void main(String[] args){
        File imgFile = new File("sample.png");
        BufferedImage img;
        try{
            img = ImageIO.read(imgFile);

            File csv = getCSV(img);
        }catch(IOException e){
            e.printStackTrace();
        }
    }
    public static File getCSV(BufferedImage img, File outFile){
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
        return outFile;
    }

    public static File getCSV(BufferedImage img){
        File result = new File("chart.csv");
        return getCSV(img, result);
    }



}
