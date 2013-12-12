package test;

import java.io.File;
import java.io.IOException;

import jxl.Workbook;
import jxl.write.WritableImage;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

public class ReadImageSample {

	public static void insertImg(WritableSheet dataSheet,	int col,int row,int width,int height,File imgFile){
		WritableImage img = new WritableImage(col,row,width,height,imgFile);
		dataSheet.addImage(img);
	}
	
	public static void main(String arvg[]){
		WritableWorkbook wwb;
		try {
			wwb = Workbook.createWorkbook(new File("E:\\img.xls"));
		
			WritableSheet imgSheet = wwb.createSheet("images", 0);
			File imgFile = new File("E:\\10698_14_110898683.jpg");
			insertImg(imgSheet, 1, 2, 4, 6, imgFile);
			wwb.write();
			wwb.close();
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
}
