package org.student.com;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

import jxl.*;
import jxl.write.*;

public class ExcelHandler {

	static final String OUTPUT = "E:\\output.xls";
	
	public void createXLS() throws Exception{
		
		WritableWorkbook wb = Workbook.createWorkbook(new File(ExcelHandler.OUTPUT));
		WritableSheet sheet = wb.createSheet("first sheet", 0);
		jxl.write.Label id = new Label(0,0,"编号");
		sheet.addCell(id);		
		jxl.write.Label sid = new Label(1,0,"学号");
		sheet.addCell(sid);
		jxl.write.Label name = new Label(2,0,"姓名");
		sheet.addCell(name);		
		jxl.write.Label sex = new Label(3,0,"性别");
		sheet.addCell(sex);		
		jxl.write.Label birth = new Label(4,0,"出生年月");
		sheet.addCell(birth);		
		jxl.write.Label college = new Label(5,0,"学院");
		sheet.addCell(college);
		wb.write();
		wb.close();
	}
	
	public void exportAsXSL(String filePath, List<StudentScore> stuScoreList ) throws Exception{
		WritableWorkbook wb = Workbook.createWorkbook(new File(filePath));
		WritableSheet sheet = wb.createSheet("first sheet", 0);
		jxl.write.Label id = new Label(0,0,"编号");
		sheet.addCell(id);		
		jxl.write.Label sid = new Label(1,0,"学号");
		sheet.addCell(sid);
		jxl.write.Label name = new Label(2,0,"姓名");
		sheet.addCell(name);		
		jxl.write.Label chinese = new Label(3,0,"语文");
		sheet.addCell(chinese);		
		jxl.write.Label math = new Label(4,0,"数学");
		sheet.addCell(math);		
		jxl.write.Label english = new Label(5,0,"英语");
		sheet.addCell(english);
		
		int row = 1;
		for(StudentScore ss : stuScoreList){
			jxl.write.Label idValue = new Label(0, row , String.valueOf(row));
			sheet.addCell(idValue);
			jxl.write.Label sidValue = new Label(1, row , ss.getId());
			sheet.addCell(sidValue);
			jxl.write.Label nameValue = new Label(2, row , ss.getName());
			sheet.addCell(nameValue);
			jxl.write.Label chinsesValue = new Label(3, row , ss.getChinese());
			sheet.addCell(chinsesValue);
			jxl.write.Label mathValue = new Label(4, row , ss.getMath());
			sheet.addCell(mathValue);
			jxl.write.Label englishValue = new Label(5, row , ss.getEnglish());
			sheet.addCell(englishValue);
			row++;
		}
		wb.write();
		wb.close();
	}
	
	public boolean isImportFile(String filePath) throws Exception{
		Workbook wb = Workbook.getWorkbook(new File(filePath));
		Sheet sheet = wb.getSheet(0);
		if(sheet.getCell(1,0).getContents().equals("学号") &&
			sheet.getCell(2,0).getContents().equals("姓名") &&
			sheet.getCell(3,0).getContents().equals("语文") &&
			sheet.getCell(4,0).getContents().equals("数学") &&
			sheet.getCell(5,0).getContents().equals("英语") )
			 	return true;
		return false;
	}
	
	public List<StudentScore> importFromXLS(String filePath) throws Exception{
		List<StudentScore> stuScoreList = new ArrayList<StudentScore>();
		Workbook wb = Workbook.getWorkbook(new File(filePath));
		Sheet sheet = wb.getSheet(0);
		int i = 1;
		while(i < sheet.getRows()){
			StudentScore ss = new StudentScore();
			ss.setId(sheet.getCell(1,i).getContents());
			ss.setName(sheet.getCell(2, i).getContents());
			ss.setChinese(sheet.getCell(3, i).getContents());
			ss.setMath(sheet.getCell(4, i).getContents());
			ss.setEnglish(sheet.getCell(5, i).getContents());
			stuScoreList.add(ss);
			i++;
		}
		return stuScoreList;
	}
	
	public void deleteRow(String stuId) throws Exception {
		Workbook wb = Workbook.getWorkbook(new File(ExcelHandler.OUTPUT));
		WritableWorkbook wwb = Workbook.createWorkbook(new File(ExcelHandler.OUTPUT), wb);
		WritableSheet sheet = wwb.getSheet(0);
		int i = 1;
		while( i < sheet.getRows()){
			if(sheet.getCell(1, i).getContents().equals(stuId)){
				sheet.removeRow(i);
				break;
			}else{
				i++;
			}
		}
		wwb.write();
		wwb.close();
	}
	
	public List<Student> readXLS() throws Exception{
		List<Student> stuList = new ArrayList<Student>();
		Workbook wb = Workbook.getWorkbook(new File(ExcelHandler.OUTPUT));
		Sheet sheet = wb.getSheet(0);
		int i = 1;
		while(i < sheet.getRows()){
			Student stu = new Student();
			stu.setId(sheet.getCell(1, i).getContents());
			stu.setName(sheet.getCell(2, i).getContents());
			stu.setSex(sheet.getCell(3, i).getContents());
			stu.setBrithday(sheet.getCell(4, i).getContents());
			stu.setCollege(sheet.getCell(5, i).getContents());
			stuList.add(stu);
			i++;
		}
		return stuList;
	}
	
	public boolean writeXLS(Student stu) throws Exception{
		Workbook wb = Workbook.getWorkbook(new File(ExcelHandler.OUTPUT));
		WritableWorkbook wwb = Workbook.createWorkbook(new File(ExcelHandler.OUTPUT),wb);
		WritableSheet sheet = wwb.getSheet(0);
		int rows = sheet.getRows();
		int i = 1 ;
		while(i < rows){
			if(sheet.getCell(1, i).getContents().equals(stu.getId())){
				wwb.write();
				wwb.close();
				return false;
			}else{
				i++;
			}
		}
		
		jxl.write.Label id = new Label(0, rows, String.valueOf(rows));
		sheet.addCell(id);
		jxl.write.Label sid = new Label(1, rows, stu.getId());
		sheet.addCell(sid);
		jxl.write.Label name = new Label(2, rows, stu.getName());
		sheet.addCell(name);
		jxl.write.Label sex = new Label(3, rows, stu.getSex());
		sheet.addCell(sex);
		jxl.write.Label brith = new Label(4, rows, stu.getBrithday());
		sheet.addCell(brith);
		jxl.write.Label college = new Label(5, rows, stu.getCollege());
		sheet.addCell(college);
		wwb.write();
		wwb.close();
		return true;
	}
	
	private void editCell(int column , int row , String value)throws Exception{
		Workbook wb = Workbook.getWorkbook(new File(ExcelHandler.OUTPUT));
		WritableWorkbook wwb = Workbook.createWorkbook(new File(ExcelHandler.OUTPUT),wb);
		WritableSheet sheet = wwb.getSheet(0);
		WritableCell cell = sheet.getWritableCell(column, row);
		jxl.format.CellFormat cf = cell.getCellFormat();
		jxl.write.Label lbl = new Label(column,row,value);
		lbl.setCellFormat(cf);
		sheet.addCell(lbl);
		wwb.write();
		wwb.close();
	}
	
	public void modify(Student stu)throws Exception {
		Workbook wb = Workbook.getWorkbook(new File(ExcelHandler.OUTPUT));
		Sheet sheet = wb.getSheet(0);
		int i = 1;
		while( i < sheet.getRows()){
			if(sheet.getCell(1, i).getContents().equals(stu.getId())){
				editCell(2, i, stu.getName());
				editCell(3, i, stu.getSex());
				editCell(4, i, stu.getBrithday());
				editCell(5, i, stu.getCollege());
			}else{
				i++;
			}
		}
	}
}
