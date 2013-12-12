package org.student.com;

import java.awt.EventQueue;
import java.lang.reflect.InvocationTargetException;

import javax.swing.JFrame;

public class StudentInfoManagementTest {

	/**
	 * @param args
	 * @throws InvocationTargetException 
	 * @throws InterruptedException 
	 */
	public static void main(String[] args) throws InterruptedException, InvocationTargetException {
		// TODO Auto-generated method stub
		EventQueue.invokeAndWait(new Runnable() {
			
			@Override
			public void run() {
				// TODO Auto-generated method stub
				StudentManagement sf = null;
				try{
					sf = new StudentManagement();
				}catch(Exception e){
					e.printStackTrace();
				}
				sf.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
				sf.setVisible(true);
			}
		});
	}

}
