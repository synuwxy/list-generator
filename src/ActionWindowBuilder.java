package printForOffice;

import java.awt.Font;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.ItemEvent;
import java.awt.event.ItemListener;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.io.File;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

import javax.swing.DefaultComboBoxModel;
import javax.swing.GroupLayout;
import javax.swing.GroupLayout.Alignment;
import javax.swing.JButton;
import javax.swing.JComboBox;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JList;
import javax.swing.JOptionPane;
import javax.swing.JTextField;
import javax.swing.LayoutStyle.ComponentPlacement;
import javax.swing.UIManager;

public class ActionWindowBuilder 
{
	private JFrame f = new JFrame("目录生成器");
	private JTextField recordNum;
	private JTextField dirPath;
	
	private JFileChooser jFileChooser;
	private String directory;
	private List<String> fileList;
	private File fileDir;

	
	public void init()
	{
		
		JButton btnword = new JButton("生成Word");
		btnword.setFont(new Font("Dialog", 0, 14));
		btnword.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				if(null == fileList) {
					JOptionPane.showMessageDialog(null, "路径未选择", "错误", JOptionPane.INFORMATION_MESSAGE);
					return;
				}
				if(0 == fileList.size()) {
					JOptionPane.showMessageDialog(null, "路径下没有文件", "错误", JOptionPane.INFORMATION_MESSAGE);
					return;
				}
				WordUtil wu = new WordUtil();
				int count = wu.getXWPFDocument(directory, fileList);
				if(count == -1) {
					JOptionPane.showMessageDialog(null, "生成目录失败", "错误", JOptionPane.INFORMATION_MESSAGE);
					return;
				}
				recordNum.setText(count+"");
				JOptionPane.showMessageDialog(null, "生成目录成功，文件名: 目录.doc", "成功", JOptionPane.INFORMATION_MESSAGE);
			}
		});
		
		final JComboBox<String> comboBox = new JComboBox<String>(new DefaultComboBoxModel<String>());
		comboBox.addItemListener(new ItemListener() {
			@Override
			public void itemStateChanged(ItemEvent e) {
				
				switch(e.getStateChange()){
					case ItemEvent.SELECTED:
						fillFileList(e.getItem().toString());
						break;
					case ItemEvent.DESELECTED:
						break;
				}
			}
		});
		
		JButton btnexcel = new JButton("生成Excel");
		btnexcel.setFont(new Font("Dialog", 0, 14));
		btnexcel.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				if(null == fileList) {
					JOptionPane.showMessageDialog(null, "路径未选择", "错误", JOptionPane.INFORMATION_MESSAGE);
					return;
				}
				if(0 == fileList.size()) {
					JOptionPane.showMessageDialog(null, "路径下没有文件", "错误", JOptionPane.INFORMATION_MESSAGE);
					return;
				}
				ExcelUtil eu = new ExcelUtil();
				int count = eu.getHSSFWorkbook(directory, fileList);
				if(count == -1) {
					JOptionPane.showMessageDialog(null, "生成目录失败", "错误", JOptionPane.INFORMATION_MESSAGE);
					return;
				}
				recordNum.setText(count+"");
				JOptionPane.showMessageDialog(null, "生成目录成功，文件名: 目录.xls", "成功", JOptionPane.INFORMATION_MESSAGE);
			}
		});
		
		JButton button_5 = new JButton("清空文件夹路径");
		button_5.setFont(new Font("Dialog", 0, 14));
		button_5.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				directory = "";
				fileList = null;
				fileDir = null;
				dirPath.setText(directory);
				comboBox.removeAllItems();
				recordNum.setText("");
			}
		});
		JLabel label_5 = new JLabel("本次共复制");
		label_5.setFont(new Font("Dialog", 0, 14));
		recordNum = new JTextField();
		recordNum.setFont(new Font("宋体", Font.PLAIN, 20));
		recordNum.setColumns(10);
		
		JLabel label_6 = new JLabel("条记录");
		label_6.setFont(new Font("Dialog", 0, 14));
		
		JButton button_6 = new JButton("");
		//隐藏button
		button_6.setOpaque(false);
		button_6.setContentAreaFilled(false);
		button_6.setFont(new Font("Dialog", 0, 14));
		button_6.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				
			}
		});
		
		dirPath = new JTextField();
		dirPath.setFont(new Font("宋体", Font.PLAIN, 12));
		dirPath.setColumns(10);
		dirPath.addMouseListener(new MouseAdapter() {
			public void mouseClicked(MouseEvent e){
				// 选择文件夹
				if (null == directory || "".equals(directory) ){
					jFileChooser = new JFileChooser();
					jFileChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
				}
				else {
					jFileChooser = new JFileChooser(directory);
					jFileChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
				}
				if(jFileChooser.showOpenDialog(null) == JFileChooser.APPROVE_OPTION){
					fileDir = jFileChooser.getSelectedFile();
					directory = fileDir.getAbsolutePath();
					dirPath.setText(directory);
				}
			}
		});
		
		JButton button_7 = new JButton("");
		//隐藏button
		button_7.setOpaque(false);
		button_7.setContentAreaFilled(false);
		button_7.setFont(new Font("Dialog", 0, 14));
		button_7.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				
			}
		});
		
		JLabel label = new JLabel("选择文件夹");
		
		JLabel label_1 = new JLabel("使用说明:");
		
		JLabel label_2 = new JLabel("软件说明: 目录生成器，功能是将路径下的所有文件生成一个目录");
		
		JLabel lblNewLabel = new JLabel("1. 点击输入框选择一个文件夹");
		
		JList<Object> list = new JList<Object>();
		
		
		
		JButton button = new JButton("检测文件类型");
		button.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				if(null== fileDir || !fileDir.exists()){
					JOptionPane.showMessageDialog(null, "路径未选择", "错误", JOptionPane.INFORMATION_MESSAGE);
					return;
				}
				fileList = new ArrayList<String>();
				Set<String> fileSet = new HashSet<String>();
				File[] files = fileDir.listFiles();
				for (File file : files) {
					if(file.getName().indexOf(".") == -1){
						continue;
					}
					int index = getFinalPrintNum(file.getName());
					
					fileList.add(file.getName());
					fileSet.add(file.getName().substring(index+1));
				}
				comboBox.removeAllItems();
				comboBox.addItem("");
				for (String fileType : fileSet) {
					comboBox.addItem(fileType);
				}
			}
		});
		
		JLabel label_3 = new JLabel("2. 点击检测文件类型");
		
		JLabel label_4 = new JLabel("3. 在下拉框内选择文件类型");
		
		JLabel lbloffice = new JLabel("4. 点击按钮选择生成对应的office目录");
		
		JLabel label_7 = new JLabel("5. 生成的目录会在选择的文件夹中出现");
		
		GroupLayout groupLayout = new GroupLayout(f.getContentPane());
		groupLayout.setHorizontalGroup(
			groupLayout.createParallelGroup(Alignment.TRAILING)
				.addGroup(groupLayout.createSequentialGroup()
					.addGap(50)
					.addGroup(groupLayout.createParallelGroup(Alignment.LEADING)
						.addGroup(groupLayout.createSequentialGroup()
							.addPreferredGap(ComponentPlacement.RELATED)
							.addGroup(groupLayout.createParallelGroup(Alignment.LEADING)
								.addComponent(lblNewLabel, GroupLayout.PREFERRED_SIZE, 253, GroupLayout.PREFERRED_SIZE)
								.addComponent(label_1)
								.addComponent(label_2, GroupLayout.PREFERRED_SIZE, 368, GroupLayout.PREFERRED_SIZE)
								.addGroup(groupLayout.createSequentialGroup()
									.addComponent(btnword, GroupLayout.PREFERRED_SIZE, 101, GroupLayout.PREFERRED_SIZE)
									.addGap(37)
									.addComponent(btnexcel, GroupLayout.PREFERRED_SIZE, 101, GroupLayout.PREFERRED_SIZE)
									.addGap(18)
									.addComponent(list, GroupLayout.PREFERRED_SIZE, 1, GroupLayout.PREFERRED_SIZE)
									.addPreferredGap(ComponentPlacement.RELATED, GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
									.addComponent(button_5))
								.addGroup(groupLayout.createSequentialGroup()
									.addComponent(label)
									.addPreferredGap(ComponentPlacement.RELATED)
									.addComponent(dirPath, GroupLayout.PREFERRED_SIZE, 403, GroupLayout.PREFERRED_SIZE))
								.addGroup(Alignment.TRAILING, groupLayout.createSequentialGroup()
									.addComponent(button)
									.addGap(32)
									.addComponent(comboBox, GroupLayout.PREFERRED_SIZE, 79, GroupLayout.PREFERRED_SIZE))
								.addGroup(groupLayout.createSequentialGroup()
									.addComponent(label_5)
									.addGap(14)
									.addComponent(recordNum, GroupLayout.PREFERRED_SIZE, 48, GroupLayout.PREFERRED_SIZE)
									.addGap(18)
									.addComponent(label_6))
								.addComponent(label_3)
								.addComponent(label_4)
								.addComponent(lbloffice)
								.addComponent(label_7)))
						.addGroup(groupLayout.createSequentialGroup()
							.addGap(50)
							.addComponent(button_6)))
					.addPreferredGap(ComponentPlacement.RELATED)
					.addComponent(button_7)
					.addGap(44))
		);
		groupLayout.setVerticalGroup(
			groupLayout.createParallelGroup(Alignment.LEADING)
				.addGroup(groupLayout.createSequentialGroup()
					.addGap(43)
					.addGroup(groupLayout.createParallelGroup(Alignment.BASELINE)
						.addComponent(label)
						.addComponent(dirPath, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE, GroupLayout.PREFERRED_SIZE))
					.addGap(18)
					.addGroup(groupLayout.createParallelGroup(Alignment.BASELINE)
						.addComponent(comboBox, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE, GroupLayout.PREFERRED_SIZE)
						.addComponent(button))
					.addGap(38)
					.addComponent(label_2)
					.addGap(9)
					.addComponent(label_1)
					.addPreferredGap(ComponentPlacement.UNRELATED)
					.addComponent(lblNewLabel)
					.addPreferredGap(ComponentPlacement.RELATED)
					.addComponent(label_3)
					.addPreferredGap(ComponentPlacement.RELATED)
					.addComponent(label_4)
					.addPreferredGap(ComponentPlacement.RELATED)
					.addComponent(lbloffice)
					.addPreferredGap(ComponentPlacement.RELATED)
					.addComponent(label_7)
					.addGap(55)
					.addGroup(groupLayout.createParallelGroup(Alignment.BASELINE)
						.addComponent(btnword)
						.addComponent(btnexcel)
						.addComponent(list, GroupLayout.PREFERRED_SIZE, 1, GroupLayout.PREFERRED_SIZE)
						.addComponent(button_5))
					.addGap(18)
					.addGroup(groupLayout.createParallelGroup(Alignment.BASELINE)
						.addComponent(button_6)
						.addComponent(button_7))
					.addPreferredGap(ComponentPlacement.UNRELATED)
					.addGroup(groupLayout.createParallelGroup(Alignment.BASELINE)
						.addComponent(label_5)
						.addComponent(recordNum, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE, GroupLayout.PREFERRED_SIZE)
						.addComponent(label_6))
					.addContainerGap(66, Short.MAX_VALUE))
		);
		//设置图标
//		ImageIcon imageIcon = new ImageIcon(ActionWindowBuilder.class.getResource("/images/synuLogo.jpg"));
//		f.setIconImage(imageIcon.getImage());
		//设置布局
		f.getContentPane().setLayout(groupLayout);
		//设置大小
		f.setSize(587, 545);
		//设置打开居中
		f.setLocationRelativeTo(null);
		//设置关闭时完全关闭
		f.setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
		//设置不可随意改变大小
		f.setResizable(false);
		//设置窗口展示
		f.setVisible(true);
		 
	}
	
	
	public static void main(String[] args) {
		try {
			//设置为Windows样式
			UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
		} catch (Exception e) {
			e.printStackTrace();
		} 
		//启动
		new ActionWindowBuilder().init();
	}
	
	public int getFinalPrintNum(String fileName){
		char[] ch = fileName.toCharArray();
		int i = ch.length-1;
		for(;i >=0;i--){
			if(".".equals(ch[i]+"")){
				break;
			}
		}
		return i;
	}
	public void fillFileList (String type){
		fileList = new ArrayList<String>();
		File[] files = fileDir.listFiles();
		for (File file : files) {
			if(getFinalPrintNum(file.getName()) == -1) {
				continue;
			}
			String fileType = file.getName().substring(getFinalPrintNum(file.getName())+1);
			if(file.getName().indexOf(".") == -1){
				continue;
			}
			else if(fileType.equals("") || (!"".equals(type) && !fileType.equals(type))){
				continue;
			}
			fileList.add(file.getName().substring(0, getFinalPrintNum(file.getName())));
		}
	}
}
