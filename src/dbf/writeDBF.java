package dbf;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;

import com.linuxense.javadbf.DBFField;
import com.linuxense.javadbf.DBFWriter;

public class writeDBF {
	public static void writeDBF(String path) {
		OutputStream fos = null;
		try {
			// 定义DBF文件字段
			DBFField[] fields = new DBFField[3];
			// 分别定义各个字段信息，setFieldName和setName作用相同,
			// 只是setFieldName已经不建议使用
			fields[0] = new DBFField();
			// fields[0].setFieldName("emp_code");
			fields[0].setName("KDBH");
			fields[0].setDataType(DBFField.FIELD_TYPE_C);
			fields[0].setFieldLength(10);
			fields[1] = new DBFField();
			// fields[1].setFieldName("emp_name");
			fields[1].setName("KDMC");
			fields[1].setDataType(DBFField.FIELD_TYPE_C);
			fields[1].setFieldLength(50);
			fields[2] = new DBFField();
			// fields[2].setFieldName("salary");
			fields[2].setName("SIZE");
			fields[2].setDataType(DBFField.FIELD_TYPE_N);
			fields[2].setFieldLength(5);
			fields[2].setDecimalCount(0);
			//DBFWriter writer = new DBFWriter(new File(path));
			// 定义DBFWriter实例用来写DBF文件
			DBFWriter writer = new DBFWriter();
			writer.setCharactersetName("GBK");
			// 把字段信息写入DBFWriter实例，即定义表结构
			writer.setFields(fields);
			// 一条条的写入记录
			Object[] rowData ;
			for(int i = 1 ;i< 301 ; i++){
				rowData = new Object[3];
				if(i<10){
					rowData[0] = "1000"+i;
					rowData[1] = "中三-300"+i;
					rowData[2] = new Double(30+i%30);
				}else if(i<100){
					rowData[0] = "100"+i;
					rowData[1] = "中三-30"+i;
					rowData[2] = new Double(30+i%30);
				}else{
					rowData[0] = "10"+i;
					rowData[1] = "中三-3"+i;
					rowData[2] = new Double(30+i%30);
				}
				writer.addRecord(rowData);
				System.out.println(rowData[0]+":"+rowData[1]+":"+rowData[2]);
			}
			
			// 定义输出流，并关联的一个文件
			fos = new FileOutputStream(path);
			// 写入数据
			writer.write(fos);
			//writer.write();
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			try {
				fos.close();
			} catch (Exception e) {
			}
		}
	}
	public static void main(String[] args) {
		writeDBF("E:\\google下载\\2014准确数据aaaaaaaa\\kdk-test1.dbf");
	}
}
