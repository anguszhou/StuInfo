package test;

import java.awt.BorderLayout;

import javax.swing.JFrame;
import javax.swing.event.TableModelListener;
import javax.swing.table.TableModel;

import com.hg.xdoc.XDoc;
import com.hg.xdoc.XDocViewer;
import com.hg.xdoc.model.Body;
import com.hg.xdoc.model.Color;
import com.hg.xdoc.model.Img;
import com.hg.xdoc.model.Para;
import com.hg.xdoc.model.Rect;
import com.hg.xdoc.model.SText;
import com.hg.xdoc.model.Table;
import com.hg.xdoc.model.Text;

/**
 * TableModel打印预览
 * @author xdoc
 */
public class TableModelPrint {
    /**
     * TableModel打印预览
     * @param args
     */
    public static void main(String[] args) {
        //模拟TableModel
        TableModel model = new TableModel() {
            public int getRowCount() {
                return 100;
            }
            public int getColumnCount() {
                return 3;
            }
            public String getColumnName(int columnIndex) {
                return "第" + columnIndex + "列";
            }
            public Class getColumnClass(int columnIndex) {
                return String.class;
            }
            public boolean isCellEditable(int rowIndex, int columnIndex) {
                return false;
            }
            public Object getValueAt(int rowIndex, int columnIndex) {
                return String.valueOf(Math.random());
            }
            public void setValueAt(Object aValue, int rowIndex, int columnIndex) {}
            public void addTableModelListener(TableModelListener l) {}
            public void removeTableModelListener(TableModelListener l) {}
        };
        XDoc xdoc = new XDoc();
        
        //获取背景
        Rect back = xdoc.getBack();
        //添加LOGO徽标
        back.add(new Img("http://www.baidu.com/img/baidu_sylogo1.gif"));
        //添加水印
        SText stext = new SText("XDOC文档模型");
        stext.setColor(null);
        stext.setFillColor(Color.pink);
        stext.setBold(true);
        stext.setSize(400, 200);
        stext.setRotate(-45);
        stext.setLocation((back.getWidth() - stext.getWidth()) / 2, (back.getHeight() - stext.getHeight()) / 2);
        back.add(stext);
        //添加页码
        Rect pageRect = new Rect();
        Para para = new Para("第#pageno#页/共#pagecount#页");
        para.setAlign(Para.ALIGN_CENTER);
        pageRect.add(para);
        pageRect.setColor(null);
        pageRect.setWidth(200);
        pageRect.setLocation((back.getWidth() - pageRect.getWidth()) / 2, back.getHeight() - pageRect.getHeight() - 40);
        back.add(pageRect);

        Body body = xdoc.getBody();
        body.add(new Para("TableModel打印预览演示", 1));
        body.add(new Para());
        Table tab = new Table(model.getRowCount(), model.getColumnCount());
        tab.setSizeType(Table.SIZE_TYPE_AUTOSIZE);
        Rect rect;
        for (int i = 0; i < model.getColumnCount(); i++) {
            rect = new Rect();
            rect.setFillColor(Color.orange);
            rect.add(new Text(model.getColumnName(i)));
            rect.setSizeType(Rect.SIZE_TYPE_AUTOSIZE);
            tab.add(rect, 0, i);
        }
        Color c = new Color(Color.lightgray, 128);
        for (int i = 0; i < model.getRowCount(); i++) {
            for (int j = 0; j < model.getColumnCount(); j++) {
                rect = new Rect();
                if (i % 2 == 1) rect.setFillColor(c);
                rect.add(new Text(String.valueOf(model.getValueAt(i, j))));
                rect.setSizeType(Rect.SIZE_TYPE_AUTOSIZE);
                tab.add(rect, i + 1, j);
            }
        }
        body.add(new Para(tab));
        XDocViewer viewer = new XDocViewer();
        JFrame viewFrame = new JFrame("基于XDOC阅读器的JTable打印预览");
        viewFrame.getContentPane().add(viewer, BorderLayout.CENTER);
        viewFrame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        viewFrame.setVisible(true);
        viewFrame.setBounds(0, 0, 800, 600);
        viewFrame.setExtendedState(JFrame.MAXIMIZED_BOTH);
        viewer.open(xdoc);
    }
}
