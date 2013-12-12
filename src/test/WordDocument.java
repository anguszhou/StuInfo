package test;

import java.util.*;

import com.jacob.activeX.ActiveXComponent;//加入activeX控件包
import com.jacob.com.Dispatch; //加入接口包
import com.jacob.com.Variant;

/**
 * <p>标题: word文档</p>
 * <p>描述: word文档的定义</p>
 * <p>版权: Copyright (c) 2004</p>
 * <p>公司:</p>
 * @author wl
 * @version 1.0
 */
public class WordDocument implements java.io.Serializable{
  /**
	 * 
	 */
	private static final long serialVersionUID = 1L;

/**word应用*/
	private ActiveXComponent app;

  /**文档集对象*/
	private Object oDocs;

  /**Word运行时是否可见*/
	private boolean tVisible = false;

  /**文档对象*/
	private Object oDoc;

  /**tSaveOnExit*/
	private boolean tSaveOnExit = false;

  /**文档名称*/
	private String docName;

  /**选择区*/
	private Object oSelection;

  /**表格集*/
	private Object oTables;

  /**字体*/
  private Object oFont;

  /**段落对齐格式*/
  private Object oAlignment;

  /**页面设置*/
  Object oPageSetup;

  /**文档字体定义*/
  private WordFont font;

  /**段落样式*/
  private WordStyle style;

  /**页边距－－上*/
  private String topMargin = "2.54";

  /**页边距－－下*/
  private String bottomMargin = "2.54";

  /**页边距－－左*/
  private String leftMargin = "3.17";

  /**页边距－－右*/
  private String rightMargin = "3.17";

  /**页眉*/
  private String headerDistance = "1.5";

  /**页脚*/
  private String footerDistance = "1.75";

  	public WordDocument( throws Exception
	{
	    this.init();
	}
	
  private void init()throws Exception
  {
    this.app = new ActiveXComponent("Word.Application");
    this.app.setProperty("Visible", new Variant(tVisible));
    this.oDocs = app.getProperty("Documents").toDispatch();
    this.docName = "";
    this.font = WordFont.getInstance();
    this.style = WordStyle.getInstance();
  }

  /**
   * 获取Word版本号
   * @return String
   */
  public String getVersion()
  {
    return this.app.getProperty("version").toString();
  }

  /**
   * 判断当前系统的word是否为2000版本
   * @return boolean
   */
  public boolean isWord2000()
  {
    String ver = getVersion();
    if (ver != null && ver.length() > 2)
      ver = ver.substring(0, 3);
    if (ver.equals("9.0"))
      return true;
    else
      return false;
  }

  /**
   * 打开/新建一文档时初始对象
   * @throws Exception
   */
  private void initObj()
      throws Exception
  {
    this.oSelection = app.getProperty("Selection").toDispatch();
    this.oTables = Dispatch.get(oDoc, "Tables").toDispatch();
    this.oAlignment = Dispatch.get(this.oSelection, "ParagraphFormat").toDispatch();
    this.oFont = Dispatch.get(this.oSelection, "Font").toDispatch();
    this.oPageSetup = Dispatch.get(oDoc, "PageSetup").toDispatch();
  }

  /**
   * 新建文档
   * @param sDocName word文档名称
   * @throws Exception
   */
  public void New(String sDocName)
      throws Exception
  {
    this.docName = sDocName;
    this.oDoc = Dispatch.call(oDocs, "Add", "").toDispatch();
    this.initObj();
  }

  /**
   * 打开一文档
   * @param sDocName word文档名称
   * @throws Exception
   */
  public void Open(String sDocName)
      throws Exception
  {
    this.docName = sDocName;
    this.oDoc = Dispatch.call(oDocs, "Open", sDocName).toDispatch();
    this.initObj();
  }

  /**
   * 创建一表格
   * @param index 表格在文档的当前所有表格中的顺序号
   * @param rows 表格行数
   * @param cols 表格列数
   * @throws Exception
   * @return 表格
   */
  public WordTable createTable(int index, int rows, int cols)
      throws Exception
  {
    Object range = Dispatch.get(this.oSelection, "Range").toDispatch();
    Object oTable = Dispatch.call(oTables, "Add", range, String.valueOf(rows), String.valueOf(cols)).toDispatch();
    if (!this.isWord2000()) Dispatch.put(oTable, "Style", "网格型");

    Object item = Dispatch.call(oTables, "Item", String.valueOf(index)).toDispatch();
    WordTable table = WordTable.getInstance();

    table.setOItem(item);
    table.setOSelection(this.oSelection);
    table.setOAlignment(this.oAlignment);
    table.setOFont(this.oFont);

    return table;
  }

  /**
   * 设置字体
   * @param cWFont word字体
   * @throws Exception
   */
  public void setFont(WordFont cWFont)
      throws Exception
  {
    this.font = cWFont;
    setMyFont();
  }

  /**
   * 设置字体
   * @throws Exception
   */
  private void setMyFont()
      throws Exception
  {
    Dispatch.put(oFont, "Name", this.font.getName());
    Dispatch.put(oFont, "Size", this.font.getSize());
    if (this.font.isIsItalic()) Dispatch.put(oFont, "Italic", "1");
    else Dispatch.put(oFont, "Italic", "0");
    if (this.font.isIsBold()) Dispatch.put(oFont, "Bold", "1");
    else Dispatch.put(oFont, "Bold", "0");
    if (this.font.isIsUnderline()) Dispatch.put(oFont, "Underline", "1");
    else Dispatch.put(oFont, "Underline", "0");
  }

  /**
   * 设置段落样式
   * @param style word字体
   * @throws Exception
   */
  public void setStyle(WordStyle style)
      throws Exception
  {
    this.style = style;
    this.setMyStyle();
  }

  /**
   * 设置样式
   * @throws Exception
   */
  private void setMyStyle()
      throws Exception
  {
    if (this.style.getIAlignment() != -1)
      Dispatch.put(oAlignment, "Alignment", String.valueOf(this.style.getIAlignment()));
  }

  /**
   * 在当前光标处输入文字
   * @param text 要输入的文字
   * @throws Exception
   */
  public void write(String text)
      throws Exception
  {
    Dispatch.call(this.oSelection, "TypeText", text);
  }

  /**
   * 在当前光标处输入文字，然后换行
   * @param text String 文本
   * @throws Exception 例外
   */
  public void writeln(String text)
      throws Exception
  {
    write(text + "\r\n");
  }

  /**
   * 在当前光标处输出换行符
   * @param rowCount int 换行数
   * @throws Exception 例外
   */
  public void printEnter(int rowCount)
      throws Exception
  {
    for (int i = 0; i < rowCount; i++)
      write("\r\n");
  }

  /**
   * 换行
   * @throws Exception
   */
  public void enter()
      throws Exception
  {
    Dispatch.call(this.oSelection, "TypeParagraph");
  }

  /**
   * 光标下移一个字符
   * @throws Exception
   */
  public void moveNextChar()
      throws Exception
  {
    this.moveNextNChar(1);
  }

  /**
   * 光标下移N个字符
   * @param count 字符个数
   * @throws Exception
   */
  public void moveNextNChar(int count)
      throws Exception
  {
    Dispatch.call(this.oSelection, "MoveRight", "1", String.valueOf(count));
  }

  /**
   * 光标上移N个字符
   * @param count 字符个数
   * @throws Exception
   */
  public void moveBackNChar(int count)
      throws Exception
  {
    Dispatch.call(this.oSelection, "MoveLeft", "1", String.valueOf(count));
  }

  /**
   * 光标下移一个单词
   * @throws Exception
   */
  public void moveNextWord()
      throws Exception
  {
    this.moveNextNWord(1);
  }

  /**
   * 光标下移N个单词
   * @param count 单词个数
   * @throws Exception
   */
  public void moveNextNWord(int count)
      throws Exception
  {
    Dispatch.call(this.oSelection, "MoveRight", "2", String.valueOf(count));
  }

  /**
   * 光标上移N个单词
   * @param count 单词个数
   * @throws Exception
   */
  public void moveBackNWord(int count)
      throws Exception
  {
    Dispatch.call(this.oSelection, "MoveLeft", "2", String.valueOf(count));
  }

  /**
   * 光标下移一行
   * @throws Exception
   */
  public void moveNextLine()
      throws Exception
  {
    this.moveNextNLine(1);
  }

  /**
   * 光标下移N行
   * @param count 行数
   * @throws Exception
   */
  public void moveNextNLine(int count)
      throws Exception
  {
    Dispatch.call(this.oSelection, "MoveRight", "3", String.valueOf(count));
  }

  /**
   * 光标上移N行
   * @param count 行数
   * @throws Exception
   */
  public void moveBackNLine(int count)
      throws Exception
  {
    Dispatch.call(this.oSelection, "MoveLeft", "3", String.valueOf(count));
  }

  /**
   * 光标移动到当前行的最前
   * @throws Exception
   */
  public void moveLineHome()
      throws Exception
  {
    Dispatch.call(this.oSelection, "MoveLeft", "3", "1");
  }

  /**
   * 光标移动到当前行的最后
   * @throws Exception
   */
  public void moveLineEnd()
      throws Exception
  {
    Dispatch.call(this.oSelection, "MoveRight", "3", "1");
  }

  /**
   * 光标移动到文档最前
   * @throws Exception
   */
  public void moveDocHome()
      throws Exception
  {
    Dispatch.call(this.oSelection, "HomeKey", "6");
  }

  /**
   * 光标移动到文档最后
   * @throws Exception
   */
  public void moveDocEnd()
      throws Exception
  {
    Dispatch.call(this.oSelection, "EndKey", "6");
  }

  /**
   * 换页
   * @throws Exception
   */
  public void nextPage()
      throws Exception
  {
    Dispatch.call(this.oSelection, "InsertBreak");
  }

  /**
   * 保存文档
   * @throws Exception
   */
  public void save()
      throws Exception
  {
    this.saveAs(this.docName);
  }

  /**
   * 另存文档
   * @param  sDocName 另存文件名称
   * @throws Exception
   */
  public void saveAs(String sDocName)
      throws Exception
  {
    Dispatch.call(this.oDoc, "SaveAs", sDocName);
  }

  /**
   * 退出Word
   */
  public void close()
  {
    Dispatch.call(this.oDoc, "Close", new Variant(tSaveOnExit));
    app.invoke("Quit", new Variant[]
               {});
  }

  public WordFont getFont()
  {
    return font;
  }

  public WordStyle getStyle()
  {
    return style;
  }

  private static void test1(int mm)
  {
    WordDocument myDoc = null;
    try
    {
      myDoc = new WordDocument();
      myDoc.New("d:/work_lizq.doc");

      long d1 = new Date().getTime();
      //System.out.print("生成中.");
      WordFont myFont = myDoc.getFont();
      WordStyle myStyle = myDoc.getStyle();

      for (int i = 1; i <= mm; i++)
      {
        myFont.setIsBold(true);
        myFont.setName("仿宋_GB2312");
        myFont.setSize("15");
        myDoc.setFont(myFont);

        myStyle.setIAlignment(WordStyle.ALIGN_CENTER);
        myDoc.setStyle(myStyle);
        myDoc.write("中国建设银行审计工作底稿 —— 工作记录\r\n\r\n");

        myFont.setSize("12");
        myFont.setIsBold(false);
        myDoc.setFont(myFont);
        myStyle.setIAlignment(WordStyle.ALIGN_LEFT);
        myDoc.setStyle(myStyle);

        myDoc.write("审计项目：XXXX              总份数：XXXX               索引号：XXXX\r\n");
        myDoc.write("编 制 人：王林            编制日期：2008/07/19       共X页  第X页\r\n");

        WordTable myTable = myDoc.createTable(i, 12, 2);
        myTable.setColWidth(1, 70);
        myTable.setColWidth(2, 350);
        Thread.sleep(600);

        myStyle.setIAlignment(WordStyle.ALIGN_CENTER);
        myStyle.setIVerticalAlignment(WordStyle.VALIGN_CENTER);
        myTable.setColStyle(1, myStyle);

        myTable.splitCell(6, 1, 1, 2);
        myTable.splitCell(6, 2, 5, 1);
        myTable.splitCell(6, 3, 5, 1);

        myTable.moveToCell(1, 1);
        myDoc.write("审计分项");
        myTable.moveToCell(2, 1);
        myDoc.write("测试目标");
        myTable.moveToCell(3, 1);
        myDoc.write("取证对象");
        myTable.moveToCell(4, 1);
        myDoc.write("样本/总体");
        myTable.moveToCell(5, 1);
        myDoc.write("测\r\n试\r\n过\r\n程");
        myTable.moveToCell(6, 1);
        myDoc.write("审\r\n计\r\n发\r\n现");
        myTable.moveToCell(6, 2);
        myDoc.write("状\r\n况");
        myTable.moveToCell(7, 2);
        myDoc.write("依\r\n据");
        myTable.moveToCell(8, 2);
        myDoc.write("差\r\n异");
        myTable.moveToCell(9, 2);
        myDoc.write("影\r\n响");
        myTable.moveToCell(10, 2);
        myDoc.write("原\r\n因");
        myTable.moveToCell(11, 1);
        myDoc.write("审计\r\n建议");
        myTable.moveToCell(12, 1);
        myDoc.write("情况记录\r\n索引");
        myTable.moveToCell(13, 1);
        myDoc.write("备注");
        myTable.moveToCell(14, 1);
        myDoc.write("附件");
        myTable.moveToCell(15, 1);
        myDoc.write("主审人\r\n复核意见");
        myTable.moveToCell(15, 2);
        myDoc.write("\r\n                         签名：       XXXX年XX月XX日");
        myTable.moveToCell(16, 1);
        myDoc.write("审计组长\r\n复核意见");
        myTable.moveToCell(16, 2);
        myDoc.write("\r\n                         签名：       XXXX年XX月XX日");

        myDoc.moveDocEnd();
        if (i < (mm))
          myDoc.nextPage();

          //System.out.print(i);

      }
      myDoc.save();
      long d2 = new Date().getTime();
      //System.out.println("Time:"+((d2-d1))+"mS");
    }
    catch (Exception e)
    {
      e.printStackTrace();
    }
    finally
    {
      myDoc.close();
    }
  }

  public static void main(String[] args)
  {
    int tt = Integer.parseInt(args[0]);
    test1(tt);
  }

  public String getBottomMargin()
  {
    return bottomMargin;
  }

  public void setBottomMargin(String bottomMargin)
  {
    if (bottomMargin == null)return;

    this.bottomMargin = bottomMargin;
    Dispatch.put(oPageSetup, "BottomMargin", bottomMargin);
  }

  public String getFooterDistance()
  {
    return footerDistance;
  }

  public void setFooterDistance(String footerDistance)
  {
    if (footerDistance == null)return;
    this.footerDistance = footerDistance;
    Dispatch.put(oPageSetup, "FooterDistance", footerDistance);
  }

  public String getHeaderDistance()
  {
    return headerDistance;
  }

  public void setHeaderDistance(String headerDistance)
  {
    if (headerDistance == null)return;
    this.headerDistance = headerDistance;
    Dispatch.put(oPageSetup, "HeaderDistance", headerDistance);
  }

  public String getLeftMargin()
  {
    return leftMargin;
  }

  public void setLeftMargin(String leftMargin)
  {
    if (leftMargin == null)return;
    this.leftMargin = leftMargin;
    Dispatch.put(oPageSetup, "LeftMargin", leftMargin);
  }

  public String getRightMargin()
  {
    return rightMargin;
  }

  public void setRightMargin(String rightMargin)
  {
    if (rightMargin == null)return;
    this.rightMargin = rightMargin;
    Dispatch.put(oPageSetup, "RightMargin", rightMargin);
  }

  public String getTopMargin()
  {
    return topMargin;
  }

  public void setTopMargin(String topMargin)
  {
    if (topMargin == null)return;
    this.topMargin = topMargin;
    Dispatch.put(oPageSetup, "TopMargin", topMargin);
  }



package audit.pub.word;

/**
 * <p>标题: word字体</p>
 * <p>描述: word字体定义</p>
 * <p>版权: Copyright (c) 2004</p>
 * <p>公司:</p>
 * @author lizq
 * @version 1.0
 */
public class WordFont implements java.io.Serializable
{
  /**字体*/
  private String name = "宋体";

  /**字号*/
  private String size = "15";

  /**是否倾斜*/
  private boolean isItalic = false;

  /**是否加粗*/
  private boolean isBold = false;

  /**是否下划线*/
  private boolean isUnderline = false;

  private WordFont()
  {}

  public static WordFont getInstance()
  {
    return new WordFont();
  }

  public boolean isIsBold()
  {
    return isBold;
  }

  public boolean isIsItalic()
  {
    return isItalic;
  }

  public boolean isIsUnderline()
  {
    return isUnderline;
  }

  public String getName()
  {
    return name;
  }

  public void setName(String name)
  {
    this.name = name;
  }

  public void setIsUnderline(boolean isUnderline)
  {
    this.isUnderline = isUnderline;
  }

  public void setIsItalic(boolean isItalic)
  {
    this.isItalic = isItalic;
  }

  public void setIsBold(boolean isBold)
  {
    this.isBold = isBold;
  }

  public String getSize()
  {
    return size;
  }

  public void setSize(String size)
  {
    this.size = size;
  }
}

==============================================================================

=========================================================

==================================================================================

package audit.pub.word;

/**
 * <p>标题: word样式</p>
 * <p>描述: word样式定义</p>
 * <p>版权: Copyright (c) 2004</p>
 * <p>公司: 泰利特青岛软件研究所</p>
 * @author lizq
 * @version 1.0
 */
public class WordStyle implements java.io.Serializable
{
  /**对齐方式（水平居中）*/
  public static final int ALIGN_CENTER = 1;

  /**对齐方式（垂直 居中）*/
  public static final int VALIGN_CENTER = 1;

  /**对齐方式（右对齐）*/
  public static final int ALIGN_RIGHT = 2;

  /**对齐方式（水平左对齐）*/
  public static final int ALIGN_LEFT = 3;

  /**对齐方式（垂直下对齐）*/
  public static final int VALIGN_LEFT = 3;

  /**对齐方式（自适应）*/
  public static final int ALIGN_AUTO = 4;

  /**水平对齐方式：默认为左对齐*/
  private int iAlignment = -1;

  /**垂直对齐方式：默认为上对齐*/
  private int iVerticalAlignment = -1;

  private WordStyle()
  {}

  public static WordStyle getInstance()
  {
    return new WordStyle();
  }

  public int getIAlignment()
  {
    return iAlignment;
  }

  public void setIAlignment(int iAlignment)
  {
    this.iAlignment = iAlignment;
  }

  public int getIVerticalAlignment()
  {
    return iVerticalAlignment;
  }

  public void setIVerticalAlignment(int iVerticalAlignment)
  {
    this.iVerticalAlignment = iVerticalAlignment;
  }
}

==========================================================================

=============================================================

==========================================================================

package audit.pub.word;

import com.jacob.com.*;

/**
 * <p>标题: word表格</p>
 * <p>描述: word表格定义</p>
 * <p>版权: Copyright (c) 2004</p>
 * <p>公司: </p>
 * @author lizq
 * @version 1.0
 */
public class WordTable implements java.io.Serializable
{
  /**选择区*/
  private Object oSelection;

  /**段落对齐格式*/
  private Object oAlignment;

  /**字体*/
  private Object oFont;

  /**item*/
  private Object oItem;

//  /**行数*/
//  private int iRows = 0;
//  /**列数*/
//  private int iCols = 0;

  private WordTable()
  {}

  protected static WordTable getInstance()
  {
    return new WordTable();
  }

  /**
   * 设置表格某一行的高度
   * @param index 行序号
   * @param height 高度
   * @throws Exception
   */
  private void setRowHeight(int index, int height)
      throws Exception
  {
//     if((index<0)||(index>this.iRows)) return;
    Object oRows = Dispatch.call(oItem, "Rows", String.valueOf(index)).toDispatch();
    Dispatch.put(oRows, "Height", String.valueOf(height));
  }

  /**
   * 设置表格某一列的宽度
   * @param index 列序号
   * @param width 宽度
   * @throws Exception
   */
  public void setColWidth(int index, int width)
      throws Exception
  {
//     if((index<0)||(index>this.iCols)) return;
    Object Columns = Dispatch.call(oItem, "Columns", String.valueOf(index)).toDispatch();
    Dispatch.put(Columns, "PreferredWidth", String.valueOf(width));

  }

  /**
   * 设置表格的某一列的对齐方式
   * @param index 列序号
   * @param style 样式
   * @throws Exception
   */
  public void setColStyle(int index, WordStyle style)
      throws Exception
  {
//     Object Columns = Dispatch.call(oItem,"Columns",String.valueOf(index)).toDispatch();
//     Dispatch.call(Columns,"Select");
    this.moveToCell(1, index);
    Dispatch.call(this.oSelection, "SelectColumn");

    Object cells = Dispatch.get(this.oSelection, "Cells").toDispatch();

    int iAlignment = style.getIAlignment();
    int iVerticalAlignment = style.getIVerticalAlignment();

    if (iAlignment != -1)
      Dispatch.put(this.oAlignment, "Alignment", String.valueOf(iAlignment));
    if (iVerticalAlignment != -1)
      Dispatch.put(cells, "VerticalAlignment", String.valueOf(iVerticalAlignment));
  }

  /**
   * 设置表格的某一列的字体
   * @param index 列序号
   * @param font 字体
   * @throws Exception
   */
  public void setColFont(int index, WordFont font)
      throws Exception
  {
    this.moveToCell(1, index);
    Dispatch.call(this.oSelection, "SelectColumn");

    Dispatch.put(oFont, "Name", font.getName());
    Dispatch.put(oFont, "Size", font.getSize());
    if (font.isIsItalic()) Dispatch.put(oFont, "Italic", "1");
    else Dispatch.put(oFont, "Italic", "0");
    if (font.isIsBold()) Dispatch.put(oFont, "Bold", "1");
    else Dispatch.put(oFont, "Bold", "0");
    if (font.isIsUnderline()) Dispatch.put(oFont, "Underline", "1");
    else Dispatch.put(oFont, "Underline", "0");
  }

  /**
   * 拆分表格的某一单元格
   * @param xCell 行号
   * @param yCell 列号
   * @param iRow 拆分为几行
   * @param iCol 拆分为几列
   * @throws Exception
   */
  public void splitCell(int xCell, int yCell, int iRow, int iCol)
      throws Exception
  {
//     if((xCell<0)||(xCell>this.iRows)) return;
//     if((yCell<0)||(yCell>this.iCols)) return;

    Object cell = Dispatch.call(oItem, "Cell", String.valueOf(xCell), String.valueOf(yCell)).toDispatch();
    Dispatch.call(cell, "Split", String.valueOf(iRow), String.valueOf(iCol));
  }

  /**
   * 将光标移到表格的某一单元格上
   * @param xCell 行号
   * @param yCell 列号
   * @throws Exception
   */
  public void moveToCell(int xCell, int yCell)
      throws Exception
  {
//     if((xCell<0)||(xCell>this.iRows)) return;
//     if((yCell<0)||(yCell>this.iCols)) return;

    Object cell = Dispatch.call(oItem, "Cell", String.valueOf(xCell), String.valueOf(yCell)).toDispatch();
    Dispatch.call(cell, "Select");
  }

//  public int getICols() {
//    return iCols;
//  }
//  public void setICols(int iCols) {
//    this.iCols = iCols;
//  }
//  public int getIRows() {
//    return iRows;
//  }
//  public void setIRows(int iRows) {
//    this.iRows = iRows;
//  }
  public Object getOItem()
  {
    return oItem;
  }

  public void setOItem(Object oItem)
  {
    this.oItem = oItem;
  }

  public Object getOSelection()
  {
    return oSelection;
  }

  public void setOSelection(Object oSelection)
  {
    this.oSelection = oSelection;
  }

  public Object getOAlignment()
  {
    return oAlignment;
  }

  public void setOAlignment(Object oAlignment)
  {
    this.oAlignment = oAlignment;
  }

  public Object getOFont()
  {
    return oFont;
  }

  public void setOFont(Object oFont)
  {
    this.oFont = oFont;
  }
}
