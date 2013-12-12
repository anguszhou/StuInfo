package dbf;

import java.io.FileInputStream;
import java.io.InputStream;

import com.linuxense.javadbf.DBFField;
import com.linuxense.javadbf.DBFReader;

public class readDBF {
	public static void readDBF(String path) {
		InputStream fis = null;
		try {
			// 读取文件的输入流
			fis = new FileInputStream(path);
			// 根据输入流初始化一个DBFReader实例，用来读取DBF文件信息
			DBFReader reader = new DBFReader(fis);
			reader.setCharactersetName("GBK");
			// 调用DBFReader对实例方法得到path文件中字段的个数
			int fieldsCount = reader.getFieldCount();
			// 取出字段信息
			for (int i = 0; i < fieldsCount; i++) {
				DBFField field = reader.getField(i);
				System.out.println(field.getName());
			}
			System.out.println(reader.getRecordCount());
			//System.exit(0);
			Object[] rowValues;
			// 一条条取出path文件中记录
			while ((rowValues = reader.nextRecord()) != null) {
				for (int i = 0; i < rowValues.length; i++) {
					System.out.println(rowValues[i]);
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			try {
				fis.close();
			} catch (Exception e) {
			}
		}
	}
	
	public static void main(String[] args) {
		readDBF.readDBF("E:\\google下载\\2014准确数据aaaaaaaa\\10698sbk（上报库）.dbf");
	}
}

