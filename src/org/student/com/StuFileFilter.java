package org.student.com;

import java.io.File;

import javax.swing.filechooser.FileFilter;

public class StuFileFilter extends FileFilter {

	String ext;
	
	public StuFileFilter(String ext){
		this.ext = ext;
	}
	
	@Override
	public boolean accept(File file) {
		// TODO Auto-generated method stub
		if(file.isDirectory())
			return true;
		
		String filename = file.getName();
		int index = filename.lastIndexOf('.');
		if( index > 0 && index < filename.length() - 1){
			String extension = filename.substring(index+1).toLowerCase();
			if(extension.equals(ext)){
				return true;
			}
		}
		return false;
	}

	@Override
	public String getDescription() {
		// TODO Auto-generated method stub
		if(ext.equals("xls")){
			return "Office Excel File(*.xls)";
		}
		if(ext.equals("xlsx")){
			return "Office2007 Excel File(*.xlsx)";
		}
		return "";
	}

}
