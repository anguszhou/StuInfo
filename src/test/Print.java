package test;

import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.Font;
import java.awt.Graphics;
import java.awt.Graphics2D;
import java.awt.PrintJob;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.print.PageFormat;
import java.awt.print.Printable;
import java.awt.print.PrinterException;
import java.awt.print.PrinterJob;
import java.io.File;
import java.io.FileInputStream;
import java.util.Properties;

import javax.print.Doc;
import javax.print.DocFlavor;
import javax.print.DocPrintJob;
import javax.print.PrintException;
import javax.print.PrintService;
import javax.print.PrintServiceLookup;
import javax.print.ServiceUI;
import javax.print.SimpleDoc;
import javax.print.attribute.DocAttributeSet;
import javax.print.attribute.HashDocAttributeSet;
import javax.print.attribute.HashPrintRequestAttributeSet;
import javax.print.attribute.PrintRequestAttributeSet;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTextArea;

import org.student.com.StudentManagement;

public class Print extends JFrame
implements ActionListener, Printable
{
	private JButton printTextButton = new JButton("Print Text");
    private JButton previewButton = new JButton("Print Preview");
	private JButton printText2Button = new JButton("Print Text2");
	private JButton printFileButton = new JButton("Print File");
    private JButton printFrameButton = new JButton("Print Frame");
    private JButton exitButton = new JButton("Exit");
    private JLabel tipLabel = new JLabel("");
    private JTextArea area = new JTextArea();
    private JScrollPane scroll = new JScrollPane(area);
    private JPanel buttonPanel = new JPanel();

    private int PAGES = 0;
    private String printStr="";

    public Print(String str)
    {
    	printStr=str;
        this.setTitle("Print Test");
        this.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        this.setBounds((int)((SystemProperties.SCREEN_WIDTH - 800) / 2), (int)((SystemProperties.SCREEN_HEIGHT - 600) / 2), 800, 600);
        initLayout();
    }

    private void initLayout()
    {
        this.getContentPane().setLayout(new BorderLayout());
        this.getContentPane().add(scroll, BorderLayout.CENTER);
        printTextButton.setMnemonic('P');
        printTextButton.addActionListener(this);
        buttonPanel.add(printTextButton);
        previewButton.setMnemonic('v');
        previewButton.addActionListener(this);
        buttonPanel.add(previewButton);
        printText2Button.setMnemonic('e');
        printText2Button.addActionListener(this);
        buttonPanel.add(printText2Button);
        printFileButton.setMnemonic('i');
        printFileButton.addActionListener(this);
        buttonPanel.add(printFileButton);
        printFrameButton.setMnemonic('F');
        printFrameButton.addActionListener(this);
        buttonPanel.add(printFrameButton);
        exitButton.setMnemonic('x');
        exitButton.addActionListener(this);
        buttonPanel.add(exitButton);
        this.getContentPane().add(buttonPanel, BorderLayout.SOUTH);
    }

    public void actionPerformed(ActionEvent evt)
    {
        Object src = evt.getSource();
        if (src == printTextButton)
            printTextAction();
        else if (src == previewButton)
            previewAction();
        else if (src == printText2Button)
            printText2Action();
        else if (src == printFileButton)
            printFileAction();
        else if (src == printFrameButton)
            printFrameAction();
        else if (src == exitButton)
            exitApp();
    }

    public int print(Graphics g, PageFormat pf, int page) throws PrinterException
	{
        Graphics2D g2 = (Graphics2D)g;
        g2.setPaint(Color.black);
	    if (page >= PAGES)
	        return Printable.NO_SUCH_PAGE;
        g2.translate(pf.getImageableX(), pf.getImageableY());
        drawCurrentPageText(g2, pf, page);
        return Printable.PAGE_EXISTS;
	}

    private void drawCurrentPageText(Graphics2D g2, PageFormat pf, int page)
	{
        Font f = area.getFont();
        String s = getDrawText(printStr)[page];
        String drawText;
        float ascent = 16;
        int k, i = f.getSize(), lines = 0;
        while(s.length() > 0 && lines < 54)
        {
            k = s.indexOf('\n');
            if (k != -1)
            {
                lines += 1;
                drawText = s.substring(0, k);
                g2.drawString(drawText, 0, ascent);
                if (s.substring(k + 1).length() > 0)
                {
                    s = s.substring(k + 1);
                    ascent += i;
                }
            }
            else
            {
                lines += 1;
                drawText = s;
                g2.drawString(drawText, 0, ascent);
                s = "";
            }
        }
	}

    public String[] getDrawText(String s)
    {
        String[] drawText = new String[PAGES];
        for (int i = 0; i < PAGES; i++)
            drawText[i] = "";

        int k, suffix = 0, lines = 0;
        while(s.length() > 0)
        {
            if(lines < 54)
            {
                k = s.indexOf('\n');
                if (k != -1)
                {
                    lines += 1;
                    drawText[suffix] = drawText[suffix] + s.substring(0, k + 1);
                    if (s.substring(k + 1).length() > 0)
                        s = s.substring(k + 1);
                }
                else
                {
                    lines += 1;
                    drawText[suffix] = drawText[suffix] + s;
                    s = "";
                }
            }
            else
            {
                lines = 0;
                suffix++;
            }
        }
        return drawText;
    }

    public int getPagesCount(String curStr)
	{
        int page = 0;
        int position, count = 0;
        String str = curStr;
	    while(str.length() > 0)
	    {
	        position = str.indexOf('\n');
            count += 1;
	        if (position != -1)
                str = str.substring(position + 1);
	        else
	            str = "";
	    }

	    if (count > 0)
	        page = count / 54 + 1;

        return page;
	}

    private void printTextAction()
    {
        if (printStr != null && printStr.length() > 0)
        {
            PAGES = getPagesCount(printStr);
            PrinterJob myPrtJob = PrinterJob.getPrinterJob();
            PageFormat pageFormat = myPrtJob.defaultPage();
            myPrtJob.setPrintable(this, pageFormat);
            if (myPrtJob.printDialog())
            {
                try
                {
                    myPrtJob.print();
                }
                catch(PrinterException pe)
                {
                    pe.printStackTrace();
                }
            }
        }
        else
        {
            JOptionPane.showConfirmDialog(null, "Sorry, Printer Job is Empty, Print Cancelled!", "Empty"
                                        , JOptionPane.DEFAULT_OPTION, JOptionPane.WARNING_MESSAGE);
        }
    }

    private void previewAction()
    {
        PAGES = getPagesCount(printStr);
        (new PrintPreviewDialog(this, "Print Preview", true, this, printStr)).setVisible(true);
    }

    private void printText2Action()
    {
        if (printStr != null && printStr.length() > 0)
        {
            PAGES = getPagesCount(printStr);
            DocFlavor flavor = DocFlavor.SERVICE_FORMATTED.PRINTABLE;
            PrintService printService = PrintServiceLookup.lookupDefaultPrintService();
            DocPrintJob job = printService.createPrintJob();
            PrintRequestAttributeSet pras = new HashPrintRequestAttributeSet();
            DocAttributeSet das = new HashDocAttributeSet();
            Doc doc = new SimpleDoc(this, flavor, das);

            try
            {
                job.print(doc, pras);
            }
            catch(PrintException pe)
            {
                pe.printStackTrace();
            }
        }
        else
        {
            JOptionPane.showConfirmDialog(null, "Sorry, Printer Job is Empty, Print Cancelled!", "Empty"
                                        , JOptionPane.DEFAULT_OPTION, JOptionPane.WARNING_MESSAGE);
        }
    }

    private void printFileAction()
    {
        JFileChooser fileChooser = new JFileChooser(SystemProperties.USER_DIR);
        //fileChooser.setFileFilter(new com.szallcom.file.JavaFilter());
        int state = fileChooser.showOpenDialog(this);
        if (state == fileChooser.APPROVE_OPTION)
        {
            File file = fileChooser.getSelectedFile();
            PrintRequestAttributeSet pras = new HashPrintRequestAttributeSet();
            DocFlavor flavor = DocFlavor.INPUT_STREAM.AUTOSENSE;
            PrintService printService[] = PrintServiceLookup.lookupPrintServices(flavor, pras);
            PrintService defaultService = PrintServiceLookup.lookupDefaultPrintService();
            PrintService service = ServiceUI.printDialog(null, 200, 200, printService
                                                        , defaultService, flavor, pras);
            if (service != null)
            {
                try
                {
                    DocPrintJob job = service.createPrintJob();
                    FileInputStream fis = new FileInputStream(file);
                    DocAttributeSet das = new HashDocAttributeSet();
                    Doc doc = new SimpleDoc(fis, flavor, das);
                    job.print(doc, pras);
                }
                catch(Exception e)
                {
                    e.printStackTrace();
                }
            }
        }
    }

    private void printFrameAction()
    {
        Toolkit kit = Toolkit.getDefaultToolkit();
        Properties props = new Properties();
        props.put("awt.print.printer", "durango");
        props.put("awt.print.numCopies", "2");
        if(kit != null)
        {
            PrintJob printJob = kit.getPrintJob(this, "Print Frame", props);
            if(printJob != null)
            {
                Graphics pg = printJob.getGraphics();
                if(pg != null)
                {
					try
					{
                        this.printAll(pg);
                    }
                    finally
                    {
                        pg.dispose();
                    }
                }
                printJob.end();
            }
        }
    }

    private void exitApp()
    {
        this.setVisible(false);
        this.dispose();
        System.exit(0);
    }
    
    public void printTable() {
        Toolkit kit = Toolkit.getDefaultToolkit(); //获取工具箱
        Properties props = new Properties();
        props.put("awt.print.printer", "durango"); //设置打印属性
        props.put("awt.print.numCopies", "2");

        if (kit != null) {
          //获取工具箱自带的打印对象
          PrintJob printJob = kit.getPrintJob(this, "打印 页面", props);

          if (printJob != null) {
            Graphics pg = printJob.getGraphics(); //获取打印对象的图形环境
            Graphics2D g2 = (Graphics2D) pg; ///
            PageFormat pf = new PageFormat(); ///
            g2.translate(pf.getImageableX(), pf.getImageableY()); ///转换坐标，确定打印边界
            if (pg != null) {
              try {
                pg.dispose(); // Shoot the page to printer
                this.jScrollPane14.printAll(pg); //打印该窗体的组件
              }
              finally {
                pg.dispose(); //注销图形环境pageIndex
              }
            }
            printJob.end(); //结束打印作业
          }
        }
      }
    
    public static void main(String[] arvg){
    	Print sf = null;
		try{
			sf = new Print("11111111");
		}catch(Exception e){
			e.printStackTrace();
		}
		sf.setVisible(true);
    }
}
