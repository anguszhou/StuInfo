package org.student.com;

import java.awt.BorderLayout;
import java.awt.CheckboxMenuItem;
import java.awt.Color;
import java.awt.Dimension;
import java.awt.GridBagLayout;
import java.awt.Label;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.KeyAdapter;
import java.awt.event.KeyEvent;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.io.File;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.swing.AbstractAction;
import javax.swing.Action;
import javax.swing.ButtonGroup;
import javax.swing.JButton;
import javax.swing.JCheckBoxMenuItem;
import javax.swing.JComboBox;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JMenu;
import javax.swing.JMenuBar;
import javax.swing.JMenuItem;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JPopupMenu;
import javax.swing.JRadioButtonMenuItem;
import javax.swing.JScrollPane;
import javax.swing.JSeparator;
import javax.swing.JTabbedPane;
import javax.swing.JTable;
import javax.swing.JTextField;
import javax.swing.JToolBar;
import javax.swing.JTree;
import javax.swing.KeyStroke;
import javax.swing.SwingConstants;
import javax.swing.event.ChangeEvent;
import javax.swing.event.ChangeListener;
import javax.swing.event.ListSelectionEvent;
import javax.swing.event.ListSelectionListener;
import javax.swing.event.TreeModelEvent;
import javax.swing.event.TreeModelListener;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableModel;
import javax.swing.table.TableRowSorter;
import javax.swing.tree.DefaultMutableTreeNode;
import javax.swing.tree.DefaultTreeModel;
import javax.swing.tree.TreeNode;
import javax.swing.tree.TreePath;

import jxl.biff.drawing.ComboBox;
import jxl.read.biff.ExternalSheetRecord;


public class StudentManagement extends JFrame implements ActionListener,
		ChangeListener, ListSelectionListener {

	private JPanel panel;
	private JComboBox sexCom , collegeCom;
	private JTextField idTxt, nameTxt,brithTxt, sidTxt, snameTxt, chineseTxt, mathTxt, englishTxt;
	private JTree stuTree;
	private DefaultTreeModel treeModel;
	private JButton submit, cancel ,add , cel ,exportBtn, importBtn;
	private List<Student> stuList;
	private List<StudentScore> stuScoreList;
	private Student stuCurr;
	private ExcelHandler eh;
	private JCheckBoxMenuItem readItem;
	private Action saveAction, saveAsAction;
	private JPopupMenu popup;
	private JTabbedPane stuTab;
	private int tabIndex, rowIndex;
	private JTable stuTable;
	private DefaultTableModel stuDefaultTableModel;
	private JScrollPane infoSP, scoreSP;
	private JFileChooser fileChooser;
	final String output = "E:\\output.xls";
	public static final int DEDAULT_WIDTH = 800;
	public static final int DEDAULT_HEIGHT = 600;
	public static final int MAX_TABNUM = 6;
	
	public StudentManagement() throws Exception{
		setTitle("Student Information Management");
		Toolkit kit = Toolkit.getDefaultToolkit();
		Dimension dis = kit.getScreenSize();
		setSize(DEDAULT_WIDTH, DEDAULT_HEIGHT);
		
		stuList = new ArrayList<Student>();
		stuCurr = new Student();
		stuScoreList = new ArrayList<StudentScore>();
		eh = new ExcelHandler();
		File file = new File(output);
		
		try{
			if(!file.exists()){
				eh.createXLS();
			}
		}catch(Exception e){
			e.printStackTrace();
		}
		StudentPanel sp = new StudentPanel();
		add(sp);
		setMenu();
		setFileChooser();
	}
	
	private void setFileChooser(){
		fileChooser = new JFileChooser("E:\\");
		fileChooser.setFileFilter(new StuFileFilter("xls"));
	}
	
	private void setMenu() {
		// TODO Auto-generated method stub
		JMenu fileMenu = new JMenu("File");
		fileMenu.add(new MenuAction("New"));
		
		JMenuItem openItem = fileMenu.add(new MenuAction("Open"));
		openItem.setAccelerator(KeyStroke.getKeyStroke("ctrl O"));
		fileMenu.addSeparator();
		
		saveAction = new MenuAction("Save");
		JMenuItem saveItem = fileMenu.add(saveAction);
		saveItem.setAccelerator(KeyStroke.getKeyStroke("ctrl S"));
		
		saveAsAction = new MenuAction("Save As");
		fileMenu.add(saveAsAction);
		fileMenu.addSeparator();
		
		Action exitAction = new AbstractAction("Exit") {
			
			@Override
			public void actionPerformed(ActionEvent e) {
				// TODO Auto-generated method stub
				System.exit(0);
			}
		};
		exitAction.putValue(Action.SHORT_DESCRIPTION, "Exit APP");
		JMenuItem exitItem = fileMenu.add(exitAction);
		exitItem.setAccelerator(KeyStroke.getKeyStroke("ctrl Z"));
		
		readItem = new JCheckBoxMenuItem("Read Only");
		readItem.addActionListener(new ActionListener() {
			
			@Override
			public void actionPerformed(ActionEvent e) {
				// TODO Auto-generated method stub
				boolean saveOk = !readItem.isSelected();
				saveAction.setEnabled(saveOk);
				saveAsAction.setEnabled(saveOk);
			}
		});
		
		ButtonGroup group = new ButtonGroup();
		JRadioButtonMenuItem insertItem = new JRadioButtonMenuItem("Insert");
		insertItem.setSelected(true);
		JRadioButtonMenuItem overItem = new JRadioButtonMenuItem("overType");
		group.add(insertItem);
		group.add(overItem);
		JMenu option = new JMenu("Options");
		option.add(readItem);
		option.addSeparator();
		option.add(insertItem);
		option.add(overItem);
		
		Action deleteAction = new MenuAction("Delete");
		deleteAction.putValue(Action.SHORT_DESCRIPTION, "Delete Student Node");
		deleteAction.putValue(Action.MNEMONIC_KEY, new Integer('D'));
		Action editAction = new MenuAction("Modify");
		editAction.putValue(Action.SHORT_DESCRIPTION, "Edit Student Info");
		editAction.putValue(Action.MNEMONIC_KEY, new Integer('M'));
		
		JMenu editMenu = new JMenu("Edit");
		JMenuItem deleteItem = editMenu.add(deleteAction);
		deleteItem.setAccelerator(KeyStroke.getKeyStroke("ctrl D"));
		
		JMenuItem editItem = editMenu.add(editAction);
		editItem.setAccelerator(KeyStroke.getKeyStroke("ctrl E"));
		editMenu.addSeparator();
		editMenu.add(option);
		
		JMenu help = new JMenu("Help");
		help.setMnemonic('H');
		JMenuItem indexItem = new JMenuItem("Index");
		indexItem.setMnemonic('I');
		help.add(indexItem);
		
		Action aboutAction = new MenuAction("About");
		aboutAction.putValue(Action.MNEMONIC_KEY, new Integer('A'));
		help.add(aboutAction);
		
		JMenuBar menuBar = new JMenuBar();
		setJMenuBar(menuBar);
		menuBar.add(fileMenu);
		menuBar.add(editMenu);
		menuBar.add(help);
		
		popup = new JPopupMenu();
		popup.add(deleteAction);
		popup.addSeparator();
		popup.add(editAction);
		stuTree.setComponentPopupMenu(popup);
		
		JToolBar toolBar = new JToolBar();
		toolBar.add(editAction);
		toolBar.add(deleteAction);
		toolBar.addSeparator();
		toolBar.add(exitAction);
		add(toolBar,BorderLayout.NORTH);
		
	}

	class MenuAction extends AbstractAction{
		public MenuAction(String name){
			super(name);
		}

		@Override
		public void actionPerformed(ActionEvent e) {
			// TODO Auto-generated method stub
			if(e.getActionCommand().equals("Delete")){
				int select = JOptionPane.showConfirmDialog(null, "确定要删除该学生信息？",
											"警告",JOptionPane.YES_NO_OPTION,JOptionPane.QUESTION_MESSAGE);
				if(select == JOptionPane.YES_OPTION){
					deleteNode();
				}
			}else if(e.getActionCommand().equals("Modify")){
				int select = JOptionPane.showConfirmDialog(null, "确定要修改该信息？",
											"警告",JOptionPane.YES_NO_OPTION,JOptionPane.QUESTION_MESSAGE);
				if(select == JOptionPane.YES_OPTION){
					editNode();
				}
			}else{
				JOptionPane.showMessageDialog(null, "你点击了："+e.getActionCommand().toString(),
											"通知",JOptionPane.INFORMATION_MESSAGE);
			}
		}
		
	}
	
	class StudentPanel extends JPanel implements TreeModelListener{

		public StudentPanel() throws Exception{
			setLayout(new BorderLayout());
			stuTree = new JTree(buildStuInfoTree());
			treeModel = (DefaultTreeModel)stuTree.getModel();
			treeModel.addTreeModelListener(this);
			stuTree.addMouseListener(new MouseHandle());
			
			JScrollPane jp = new JScrollPane();
			jp.setViewportView(stuTree);
			add(jp , BorderLayout.WEST);
			
			panel = new JPanel();
			panel.setLayout(new GridBagLayout());
			JLabel id = new JLabel("学号：");
			idTxt = new JTextField();
			JLabel name = new JLabel("姓名：");
			nameTxt = new JTextField();
			JLabel sex = new JLabel("性别：");
			sexCom = new JComboBox(new String[]{"男","女"});
			JLabel brith = new JLabel("出生年月：");
			brithTxt = new JTextField();
			JLabel college = new JLabel("学院：");
			collegeCom = new JComboBox(new String[]{"软件学院","管理学院",
													"经济与金融学院","生命学院"});
			submit = new JButton("确定");
			submit.setMnemonic(KeyEvent.VK_ENTER);
			submit.addActionListener(new ActionListener() {
				
				@Override
				public void actionPerformed(ActionEvent e) {
					// TODO Auto-generated method stub
					if(idTxt.getText().trim().equals("") ||
							nameTxt.getText().trim().equals("")||
							brithTxt.getText().trim().equals("")){
						JOptionPane.showMessageDialog(null, "信息不能为空，请重新输入。", "警告", JOptionPane.INFORMATION_MESSAGE);
						clearWindow();
						return;
					}
					Student stu = new Student();
					stu.setId(idTxt.getText());
					stu.setName(nameTxt.getText());
					stu.setSex(String.valueOf(sexCom.getSelectedItem()));
					stu.setBrithday(brithTxt.getText());
					stu.setCollege(String.valueOf(collegeCom.getSelectedItem()));
					try{
						if(eh.writeXLS(stu)){
							JOptionPane.showMessageDialog(null, "成功添加学生"+stu.getName()+"的信息","通知", JOptionPane.INFORMATION_MESSAGE);
							addNode(stu);
							clearWindow();
						}else{
							JOptionPane.showMessageDialog(null, "该学号已存在，请查证后再输入。","通知", JOptionPane.INFORMATION_MESSAGE);
							clearWindow();
						}
					}catch(Exception ee){
						ee.printStackTrace();
					}
				}
			});
			
			cancel = new JButton("取消");
			cancel.setMnemonic(KeyEvent.VK_C);
			cancel.addActionListener(new ActionListener() {
				
				@Override
				public void actionPerformed(ActionEvent e) {
					// TODO Auto-generated method stub
					int select = JOptionPane.showConfirmDialog(null, "确定取消输入?", "警告", JOptionPane.YES_NO_OPTION,JOptionPane.QUESTION_MESSAGE);
					if(select == JOptionPane.YES_OPTION){
						clearWindow();
					}
				}
			});
			
			panel.add(id, new GBC(0,0).setAnchor(GBC.EAST));
			panel.add(idTxt,new GBC(1,0).setFill(GBC.HORIZONTAL).setWeight(100, 0).setInsets(1));
			panel.add(name, new GBC(0,1).setAnchor(GBC.EAST));
			panel.add(nameTxt,new GBC(1,1).setFill(GBC.HORIZONTAL).setWeight(100, 0).setInsets(1));
			panel.add(sex, new GBC(0,2).setAnchor(GBC.EAST));
			panel.add(sexCom,new GBC(1,2).setFill(GBC.HORIZONTAL).setWeight(100, 0).setInsets(1));
			panel.add(brith, new GBC(0,3).setAnchor(GBC.EAST));
			panel.add(brithTxt,new GBC(1,3).setFill(GBC.HORIZONTAL).setWeight(100, 0).setInsets(1));
			panel.add(college, new GBC(0,4).setAnchor(GBC.EAST));
			panel.add(collegeCom,new GBC(1,4).setFill(GBC.HORIZONTAL).setWeight(100, 0).setInsets(1));
			panel.add(submit, new GBC(0,5).setAnchor(GBC.EAST));
			panel.add(cancel, new GBC(1,5).setAnchor(GBC.WEST));
			
			infoSP = new JScrollPane();
			infoSP.setViewportView(panel);
			setTab();
			add(stuTab, BorderLayout.CENTER);
		}
		
		class MouseHandle extends MouseAdapter{
			public void mousePressed(MouseEvent e){
				stuTree.setEditable(false);
			}
		}

		public void addNode(Student stu){
		
			EditableNode parentNode = null;
			EditableNode newNode = new EditableNode(stu.getName(), true);
			EditableNode idleafNode = new EditableNode(stu.getId(), false);
			EditableNode sexleafNode = new EditableNode(stu.getSex(), true);
			EditableNode brithleafNode = new EditableNode(stu.getBrithday(), true);
			EditableNode collegeleafNode = new EditableNode(stu.getCollege(), true);
			newNode.add(idleafNode);
			newNode.add(sexleafNode);
			newNode.add(brithleafNode);
			newNode.add(collegeleafNode);
			newNode.setAllowsChildren(true);
			TreePath path = stuTree.getSelectionPath();
			parentNode = (EditableNode) treeModel.getRoot();
			treeModel.insertNodeInto(newNode, parentNode, parentNode.getChildCount());
			stuTree.scrollPathToVisible(new TreePath(newNode.getPath()));
		}
		
		@Override
		public void treeNodesChanged(TreeModelEvent e) {
			// TODO Auto-generated method stub
			stuTree.setEditable(false);
			TreePath tp = e.getTreePath();
			System.out.println(tp);
			EditableNode node = (EditableNode) tp.getLastPathComponent();
			
			int []index = e.getChildIndices();
			node = (EditableNode) node.getChildAt(index[0]);
			if(node.isLeaf()){
				Student stu = new Student();
				EditableNode parent = (EditableNode) node.getParent();
				stu.setId(parent.getChildAt(0).toString());
				stu.setName(parent.getUserObject().toString());
				stu.setSex(parent.getChildAt(1).toString());
				stu.setBrithday(parent.getChildAt(2).toString());
				stu.setCollege(parent.getChildAt(3).toString());
				
				try{
					eh.modify(stu);
				}catch(Exception e3){
					e3.printStackTrace();
				}
			}else{
				Student stu = new Student();
				stu.setId(node.getChildAt(0).toString());
				stu.setName(node.getUserObject().toString());
				stu.setSex(node.getChildAt(1).toString());
				stu.setBrithday(node.getChildAt(2).toString());
				stu.setCollege(node.getChildAt(3).toString());
				
				try{
					eh.modify(stu);
				}catch(Exception e3){
					e3.printStackTrace();
				}
			}
			JOptionPane.showMessageDialog(null, "修改成功！",
								"通知", JOptionPane.INFORMATION_MESSAGE);
			
		}
		
		@Override
		public void treeNodesInserted(TreeModelEvent e) {}
		@Override
		public void treeNodesRemoved(TreeModelEvent e) {}
		@Override
		public void treeStructureChanged(TreeModelEvent e) {}
	}
	
	public void editNode(){
		TreePath path = stuTree.getSelectionPath();
		if(path != null){
			EditableNode selectNode = (EditableNode) path.getLastPathComponent();
			if(selectNode.isEdit()){
				stuTree.setEditable(true);
				stuTree.startEditingAtPath(path);
			}else{
				JOptionPane.showMessageDialog(null, "不能编辑所选定信息！",
						"通知", JOptionPane.INFORMATION_MESSAGE);
			}
		}else{
			JOptionPane.showMessageDialog(null, "请先左键选中一个学生！",
					"通知", JOptionPane.INFORMATION_MESSAGE);
		}
	}
	
	public void deleteNode(){
		TreePath path = stuTree.getSelectionPath();
		if(path != null){
			EditableNode selectNode = (EditableNode) path.getLastPathComponent();
			TreeNode parent = selectNode.getParent();
			if(!selectNode.isRoot() && !selectNode.isLeaf()){
				if(parent != null){
					String stuId = selectNode.getChildAt(0).toString();
					try{
						eh.deleteRow(stuId);
					}catch(Exception e){
						e.printStackTrace();
					}
					treeModel.removeNodeFromParent(selectNode);
					JOptionPane.showMessageDialog(null, "成功删除学生"+selectNode.getUserObject().toString()+"的信息！",
							"通知", JOptionPane.INFORMATION_MESSAGE);
				}
			}else{
				JOptionPane.showMessageDialog(null, "不能删除根节点和某项学生信息！",
						"警告", JOptionPane.INFORMATION_MESSAGE);
			}
		}else{
			JOptionPane.showMessageDialog(null, "请先左键选中一个学生！",
					"通知", JOptionPane.INFORMATION_MESSAGE);
		}
	}
	
	public void clearWindow(){
		idTxt.setText("");
		nameTxt.setText("");
		sexCom.setSelectedIndex(0);
		brithTxt.setText("");
		collegeCom.setSelectedIndex(0);
	}
	
	public void clearScore(){
		sidTxt.setText("");
		snameTxt.setText("");
		chineseTxt.setText("");
		mathTxt.setText("");
		englishTxt.setText("");
	}
	
	public void addScoreToTable(StudentScore ss){
		Object[] newRow = {Integer.valueOf(ss.getId()), ss.getName(),
							Integer.valueOf(ss.getChinese()), Integer.valueOf(ss.getMath()),
							Integer.valueOf(ss.getEnglish())};
		stuDefaultTableModel.addRow(newRow);
	}
	
	public DefaultMutableTreeNode buildStuInfoTree() throws Exception {
		ExcelHandler eh = new ExcelHandler();
		List<Student> stuList = eh.readXLS();
		EditableNode root = new EditableNode("学生信息表", false);
		for(Student stu : stuList){
			EditableNode node = new EditableNode(stu.getName(), true);
			EditableNode idleafNode = new EditableNode(stu.getId(), false);
			EditableNode sexleafNode = new EditableNode(stu.getSex(), true);
			EditableNode brithleafNode = new EditableNode(stu.getBrithday(), true);
			EditableNode collegeleafNode = new EditableNode(stu.getCollege(), true);
			node.add(idleafNode);
			node.add(sexleafNode);
			node.add(brithleafNode);
			node.add(collegeleafNode);
			root.add(node);
		}
		return root;
	}
	
	public void setScoreTable(){
		String[] name = {"学号","姓名","语文","数学","英语"};
		Object[][] data ={
			{1001 ,"赵一",65,66,80},
			{1002 ,"钱二",44,98,36},
			{1003 ,"李三",99,44,55},
			{1004 ,"孙四",34,78,99},
			{1005 ,"吴无",33,89,77}
		};
		stuDefaultTableModel = new DefaultTableModel(data, name){
			public boolean isCellEditable(int row, int column){
				return false;
			}
		};
		
		stuTable = new JTable(stuDefaultTableModel);
		stuTable.getSelectionModel().addListSelectionListener(this);
		
		TableRowSorter<TableModel> stuTableSorter = 
										new TableRowSorter<TableModel>(stuTable.getModel());
		stuTableSorter.setComparator(0, new TableSort());
		stuTableSorter.setComparator(1, null);
		stuTableSorter.setComparator(2, new TableSort());
		stuTableSorter.setComparator(3, new TableSort());
		stuTableSorter.setComparator(4, new TableSort());
		stuTable.setRowSorter(stuTableSorter);
		stuTable.addMouseListener(new MouseAdapter() {
			public void mouseClicked(MouseEvent e){
				if(e.getButton() == MouseEvent.BUTTON1 && e.getClickCount() == 2){
					int[] rows = stuTable.getSelectedRows();
					int[] columns = stuTable.getSelectedColumns();
					String temp = String.valueOf(stuTable.getValueAt(rows[0], columns[0]));
					JOptionPane.showMessageDialog(null, "双击选中的内容是："+temp,
							"通知", JOptionPane.INFORMATION_MESSAGE);
				}
			}
		});
		stuTable.addKeyListener(new KeyAdapter() {
			public void keyTyped(KeyEvent e){
				char c = e.getKeyChar();
				JOptionPane.showMessageDialog(null, "你在键盘上按下了："+c,
						"通知", JOptionPane.INFORMATION_MESSAGE);
			}
		});
	}
	
	public void setScoreSP(){
		rowIndex = -1;
		JPanel input = new JPanel();
		input.setLayout(new GridBagLayout());
		JLabel sid = new JLabel("学号：");
		sidTxt = new JTextField();
		JLabel sname = new JLabel("姓名：");
		snameTxt = new JTextField();
		JLabel chinese= new JLabel("语文：");
		chineseTxt = new JTextField();
		JLabel math = new JLabel("数学：");
		mathTxt = new JTextField();
		JLabel english = new JLabel("英语：");
		englishTxt = new JTextField();
		add = new JButton("添加/修改");
		add.setMnemonic(KeyEvent.VK_ENTER);
		add.addActionListener(this);
		cel = new JButton("取消");
		cel.setMnemonic(KeyEvent.VK_C);
		cel.addActionListener(this);
		
		exportBtn = new JButton("导出表格");
		exportBtn.setMnemonic(KeyEvent.VK_O);
		exportBtn.addActionListener(this);
		importBtn = new JButton("导入表格");
		importBtn.setMnemonic(KeyEvent.VK_I);
		importBtn.addActionListener(this);
		
		input.add(sid, new GBC(0,0).setAnchor(GBC.EAST));
		input.add(sidTxt , new GBC(1,0,5,1).setFill(GBC.HORIZONTAL).setWeight(100, 0).setInsets(1));
		input.add(sname, new GBC(0,1).setAnchor(GBC.EAST));
		input.add(snameTxt , new GBC(1,1,5,1).setFill(GBC.HORIZONTAL).setWeight(100, 0).setInsets(1));
		input.add(chinese, new GBC(0,2).setAnchor(GBC.EAST));
		input.add(chineseTxt , new GBC(1,2,5,1).setFill(GBC.HORIZONTAL).setWeight(100, 0).setInsets(1));
		input.add(math, new GBC(0,3).setAnchor(GBC.EAST));
		input.add(mathTxt , new GBC(1,3,5,1).setFill(GBC.HORIZONTAL).setWeight(100, 0).setInsets(1));
		input.add(english, new GBC(0,4).setAnchor(GBC.EAST));
		input.add(englishTxt , new GBC(1,4,5,1).setFill(GBC.HORIZONTAL).setWeight(100, 0).setInsets(1));
		input.add(add, new GBC(0,5).setAnchor(GBC.EAST));
		input.add(cel, new GBC(1,5).setAnchor(GBC.WEST));
		input.add(exportBtn, new GBC(6,5).setAnchor(GBC.EAST));
		input.add(importBtn, new GBC(7,5).setAnchor(GBC.WEST));
		
		setScoreTable();
		JScrollPane tableScore = new JScrollPane(stuTable);
		
		JPanel score = new JPanel();
		score.setLayout(new BorderLayout());
		score.add(input, BorderLayout.NORTH);
		score.add(new JSeparator(), BorderLayout.CENTER);
		score.add(tableScore, BorderLayout.SOUTH);
		
		scoreSP = new JScrollPane();
		scoreSP.setViewportView(score);
	
	}
	
	public void setTab(){
		tabIndex = 0 ;
		setScoreSP();
		stuTab = new JTabbedPane();
		stuTab.addTab("学生信息", infoSP);
		stuTab.addTab("学生成绩", scoreSP);
		stuTab.addTab("测试Tab", newJPanel());
		
		stuTab.addChangeListener(this);
	}
	
	private boolean stringIsNum(String s){
		Pattern pattern = Pattern.compile("[0-9]*");
		Matcher isNum = pattern.matcher(s);
		if(isNum.matches()){
			return true;
		}else{
			return false;
		}
	}
	
	private JPanel newJPanel(){
		JPanel jp = new JPanel();
		JLabel label = new JLabel("new Label", SwingConstants.CENTER);
		label.setBackground(Color.yellow);
		label.setOpaque(true);
		jp.add(label);
		return jp;
	}
	
	@Override
	public void valueChanged(ListSelectionEvent e) {
		// TODO Auto-generated method stub
		int[] rows = stuTable.getSelectedRows();
		int[] columns = stuTable.getSelectedColumns();
		if(rows.length == 1){
			rowIndex = rows[0];
			sidTxt.setText(stuTable.getValueAt(rowIndex, 0).toString());
			snameTxt.setText(stuTable.getValueAt(rowIndex, 1).toString());
			chineseTxt.setText(stuTable.getValueAt(rowIndex, 2).toString());
			mathTxt.setText(stuTable.getValueAt(rowIndex, 3).toString());
			englishTxt.setText(stuTable.getValueAt(rowIndex, 4).toString());
		}
	}

	@Override
	public void stateChanged(ChangeEvent e) {
		// TODO Auto-generated method stub
		if(stuTab.getSelectedIndex() == stuTab.getTabCount()-1 && stuTab.getTabCount() < MAX_TABNUM){
			stuTab.addTab("test Tab"+String.valueOf(stuTab.getTabCount()-1), newJPanel());
		}else if(stuTab.getTabCount()>3 && stuTab.getSelectedIndex()==stuTab.getTabCount()-2){
			stuTab.removeTabAt(stuTab.getTabCount()-1);
		}
	}

	@Override
	public void actionPerformed(ActionEvent e) {
		// TODO Auto-generated method stub
		
		if(e.getActionCommand().equals("添加/修改")){
			if(rowIndex < 0){
				if(sidTxt.getText().trim().equals("")||
						snameTxt.getText().trim().equals("")||
						chineseTxt.getText().trim().equals("")||
						mathTxt.getText().trim().equals("")||
						englishTxt.getText().trim().equals("")){
					JOptionPane.showMessageDialog(null, "成绩信息不能为空！",
							"警告", JOptionPane.INFORMATION_MESSAGE);
					clearScore();
					return;
				}else if(!stringIsNum(sidTxt.getText().trim())||
						!stringIsNum(chineseTxt.getText().trim())||
						!stringIsNum(mathTxt.getText().trim())||
						!stringIsNum(englishTxt.getText().trim())){
					JOptionPane.showMessageDialog(null, "学号与成绩必须是数字",
							"警告", JOptionPane.INFORMATION_MESSAGE);
					clearScore();
					return;
				}else{
					StudentScore ss = new StudentScore();
					ss.setId(sidTxt.getText().trim());
					ss.setName(snameTxt.getText().trim());
					ss.setChinese(chineseTxt.getText().trim());
					ss.setMath(mathTxt.getText().trim());
					ss.setEnglish(englishTxt.getText().trim());
					addScoreToTable(ss);
					JOptionPane.showMessageDialog(null, "成功添加"+ss.getName()+"的成绩",
							"通知", JOptionPane.INFORMATION_MESSAGE);
					clearScore();
				}
			}
		}else if(e.getActionCommand().equals("取消")){
			
			int select = JOptionPane.showConfirmDialog(null, "确定取消输入？","警告",
										JOptionPane.YES_NO_OPTION,JOptionPane.QUESTION_MESSAGE);
			if(select == JOptionPane.YES_OPTION){
				clearScore();
				rowIndex = -1;
			}
			
		}else if(e.getActionCommand().equals("导出表格")){
			fileChooser.setDialogTitle("导出学生成绩信息");
			fileChooser.setSelectedFile(new File("StudentScoreExport.xls"));
			int result = fileChooser.showSaveDialog(this);
			File file = null;
			if(result == JFileChooser.APPROVE_OPTION){
				file = fileChooser.getSelectedFile();
				try{
					stuScoreList.clear();
					for (int i = 0; i < stuTable.getRowCount(); i++) {
						StudentScore stuScoreCurr = new StudentScore();
						stuScoreCurr.setId(stuTable.getValueAt(i, 0).toString());
						stuScoreCurr.setName(stuTable.getValueAt(i, 1).toString());
						stuScoreCurr.setChinese(stuTable.getValueAt(i, 2).toString());
						stuScoreCurr.setMath(stuTable.getValueAt(i, 3).toString());
						stuScoreCurr.setEnglish(stuTable.getValueAt(i, 4).toString());
						stuScoreList.add(stuScoreCurr);
					}
					eh.exportAsXSL(file.getAbsolutePath(), stuScoreList);
					stuScoreList.clear();
					JOptionPane.showMessageDialog(null, "导出数据成功",
							"通知", JOptionPane.INFORMATION_MESSAGE);
				}catch(Exception e2){
					e2.printStackTrace();
				}
			}else if(result == JFileChooser.CANCEL_OPTION){
				return;
			}
		}else if(e.getActionCommand().equals("导入表格")){
			fileChooser.setDialogTitle("导入学生成绩信息");
			fileChooser.setSelectedFile(new File("StudentScoreExport.xls"));
			int result = fileChooser.showOpenDialog(this);
			File file = null;
			if(result == JFileChooser.APPROVE_OPTION){
				file = fileChooser.getSelectedFile();
				if(!file.exists()){
					JOptionPane.showMessageDialog(null, "文件不存在",
							"通知", JOptionPane.INFORMATION_MESSAGE);
					return;
				}
				try{
					stuScoreList.clear();
					if(eh.isImportFile(file.getAbsolutePath())){
						stuScoreList = eh.importFromXLS(file.getAbsolutePath());
					}else{
						JOptionPane.showMessageDialog(null, "导入文件格式不符",
								"通知", JOptionPane.INFORMATION_MESSAGE);
						return;
					}
				}catch(Exception e3){
					e3.printStackTrace();
				}
				for(StudentScore ss :stuScoreList){
					addScoreToTable(ss);
				}
				stuScoreList.clear();
				JOptionPane.showMessageDialog(null, "导入数据成功",
						"通知", JOptionPane.INFORMATION_MESSAGE);
			}else if(result == JFileChooser.CANCEL_OPTION){
				return;
			}
		}
	}

}
