package test;

import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.text.MessageFormat;

import javax.swing.*;

import org.student.com.StudentManagement;

public class PrintManager extends JDialog implements ActionListener{

	JTable jtable;
	JButton jb1, jb2;
	JPanel jp;
	boolean flag = false;
	String header;
	String footer;

	public PrintManager(Frame owner, String title, boolean model, JTable jtable
						,String header,String footer) {
		super(owner, title, model);
		this.header=header;
		this.footer=footer;
		this.jtable = jtable;
		this.setLayout(new BorderLayout());
		JScrollPane scrollPane = new JScrollPane(jtable);
		this.add(scrollPane, BorderLayout.CENTER);
		jp = new JPanel();
		jb1 = new JButton("打印");
		//jb1.setFont(MyFont.f2);
		jb1.addActionListener(this);
		jb2 = new JButton("取消");
		//jb2.setFont(MyFont.f2);
		jb2.addActionListener(this);
		jp.add(jb1);
		jp.add(jb2);
		this.add(jp, BorderLayout.SOUTH);
	
		this.setTitle("打印预览");
		this.setSize(400, 600);
	
		int width = Toolkit.getDefaultToolkit().getScreenSize().width;
		int height = Toolkit.getDefaultToolkit().getScreenSize().height;
		this.setLocation(width / 2 - this.getWidth() / 2,
		height / 2 - this.getHeight() / 2);
		this.setDefaultCloseOperation(JDialog.DISPOSE_ON_CLOSE);
		this.setVisible(true);
	}

	public void actionPerformed(ActionEvent e) {
		//响应确定打印按钮
		if (e.getSource() == jb1) {
		try {
		//设置打印格式表头
		MessageFormat headerFormat = new MessageFormat(header);
		//设置打印格式页脚
		MessageFormat footerFormat = new MessageFormat(footer);
		// 按比例将输出缩小的打印模式
		jtable.print(JTable.PrintMode.FIT_WIDTH, headerFormat,footerFormat);
		flag = true;
		} catch (Exception pe) {
		flag = false;
		System.err.println("Error printing: " + pe.getMessage());
		}
		this.dispose();
		}
		if (e.getSource() == jb2) {
		flag = false;
		this.dispose();
		}
		}
		public boolean getStatus() {
		return flag;
	}
		
	public static void main(String[] arvg){
		PrintManager sf = null;
		try{
			sf = new PrintManager();
		}catch(Exception e){
			e.printStackTrace();
		}
		sf.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		sf.setVisible(true);
	}
}
