package sd;
import java.awt.BorderLayout;
import java.awt.EventQueue;
import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.Insets;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import javax.swing.JOptionPane;
import javax.swing.*;
import javax.swing.border.EmptyBorder;

import java.io.File;

import javax.swing.filechooser.FileFilter;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFCell;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.File;
import java.util.*;
import java.io.*;
public class test extends JFrame {

	private JPanel contentPane;
	private JTextField textField;
	int StudentNumber = -1;
	int TotalScore = -1;
	int num = 0;
	boolean findNumber = false;
	boolean findScore = false;
	String path;
	String[][] tmpTableData;
	String[] book = {"流水編號","入學年度","入學方式","總成績"};
	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					test frame = new test();           
					frame.setVisible(true);
					frame.setTitle("Score");
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the frame.
	 */
	public test() {
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		contentPane = new JPanel();
		contentPane.setLayout(null);
		setContentPane(contentPane);
		test3();
	}
	void test3()
	{
		setBounds(100, 100, 450, 151);
		contentPane.setLayout(null);
		JButton btnNewButton = new JButton("瀏覽");
		btnNewButton.setBounds(355, 35, 69, 23);
		contentPane.add(btnNewButton);
		
		textField = new JTextField();
		textField.setBounds(10, 36, 335, 21);
		contentPane.add(textField);
		textField.setColumns(10);
		
		JLabel label = new JLabel("請選擇成績檔案(xls或xlsx)",0);
		label.setBounds(10, 10, 414, 15);
		contentPane.add(label);
		
		JPanel panel = new JPanel();
		panel.setBounds(10, 67, 414, 36);
		contentPane.add(panel);
		
		JButton btnNewButton_1 = new JButton("送出");
		panel.add(btnNewButton_1);
		btnNewButton.addMouseListener(new MouseAdapter() {
			public void mouseClicked(MouseEvent arg0) {
				try
				{
					JFileChooser chooser =new JFileChooser();
					chooser.setFileFilter(new MyFileFilter("xls","xlsx","請選擇要開啟的成績檔(xls或xlsx)"));
					int ret=chooser.showOpenDialog(null);
					if(ret==JFileChooser.APPROVE_OPTION)
					{
						textField.setText(chooser.getSelectedFile().getPath());
					}
				}catch(Exception e){
					System.out.println(e.toString());
				}
			}
		});
		btnNewButton_1.addMouseListener(new MouseAdapter(){
			public void mouseClicked(MouseEvent arg0){
				path = textField.getText();
				try{
				
				FileInputStream file = new FileInputStream(new File(path));
				if(path.substring(path.length()-3).equals("xls") || path.substring(path.length()-3).equals("XLS")){
					HSSFWorkbook workbook = new HSSFWorkbook(file);
		            HSSFSheet sheet = workbook.getSheetAt(0);
		            Iterator<Row> rowIterator = sheet.iterator();
		            Row row = rowIterator.next();
		            Iterator<Cell> cellIterator = row.cellIterator();
		            while(cellIterator.hasNext())
		            {
		            	Cell cell = cellIterator.next();
		            	if(cell.getCellType() == XSSFCell.CELL_TYPE_STRING)
		            	{
		            		String value = cell.getStringCellValue();
		            		System.out.println(value);
		            		int len = value.length();
		            		if(!findNumber && len >= 2)
		            		{
		            			for(int i = 0; i < len-1;i++)
		            			{
		            				if(value.charAt(i) == '學' && value.charAt(i+1) == '號')
		            				{
		            					findNumber = true;
		            					StudentNumber = cell.getColumnIndex();
		            				}
		            			}
		            		}
		            		if(!findScore && len >= 3)
		            		{
		            			for(int i = 0; i < len-2;i++)
		            			{
		            				if(value.charAt(i) == '總' && value.charAt(i+1) == '成' && value.charAt(i+2) == '績')
		            				{
		            					findScore = true;
		            					TotalScore = cell.getColumnIndex();
		            				}
		            			}
		            		}
		            	}
		            }
		            file.close();
		            workbook.close();
				}else if(path.substring(path.length()-4).equals("xlsx") || path.substring(path.length()-4).equals("XLSX")){
					XSSFWorkbook workbook = new XSSFWorkbook(file);
		            XSSFSheet sheet = workbook.getSheetAt(0);
		            Iterator<Row> rowIterator = sheet.iterator();
		            Row row = rowIterator.next();
		            Iterator<Cell> cellIterator = row.cellIterator();
		            while(cellIterator.hasNext())
		            {
		            	Cell cell = cellIterator.next();
		            	if(cell.getCellType() == XSSFCell.CELL_TYPE_STRING)
		            	{
		            		String value = cell.getStringCellValue();
		            		int len = value.length();
		            		if(!findNumber && len >= 2)
		            		{
		            			for(int i = 0; i < len-1;i++)
		            			{
		            				if(value.charAt(i) == '學' && value.charAt(i+1) == '號')
		            				{
		            					findNumber = true;
		            					StudentNumber = cell.getColumnIndex();
		            				}
		            			}
		            		}
		            		if(!findScore && len >= 3)
		            		{
		            			for(int i = 0; i < len-2;i++)
		            			{
		            				if(value.charAt(i) == '總' && value.charAt(i+1) == '成' && value.charAt(i+2) == '績')
		            				{
		            					findScore = true;
		            					TotalScore = cell.getColumnIndex();
		            				}
		            			}
		            		}
		            	}
		            }
		            file.close();
		            workbook.close();
				}}catch(Exception e){e.printStackTrace();}
				System.out.println(findNumber);
				System.out.println(findScore);
				if(findNumber && findScore)
				{
					contentPane.removeAll();
					contentPane.repaint();
					makeFinalTable();
					test2();
				}
				else
				{
					WarningDialog f = new WarningDialog("開啟的檔案不為Excel檔\n或在第一排找不到學號跟總成績");
				}
			}
		});
	}
	void makeFinalTable()
	{
		try{
			FileInputStream file = new FileInputStream(path);
			if(path.substring(path.length()-3).equals("xls") || path.substring(path.length()-3).equals("XLS")){
				HSSFWorkbook workbook = new HSSFWorkbook(file);
				HSSFSheet sheet = workbook.getSheetAt(0);
				Iterator<Row> rowIterator = sheet.iterator();
				int flow = 1;
				findNum(rowIterator);
				rowIterator = sheet.iterator();
				Row row = rowIterator.next();
				tmpTableData = new String[num][4];
				while(rowIterator.hasNext())
				{
					row = rowIterator.next();
					byte[] number = row.getCell(StudentNumber).getStringCellValue().getBytes();
					double Score = -1;
					if(isScore(row))
					{
						Score = row.getCell(TotalScore).getNumericCellValue();
					}
					String enroll = "";
					int year = (number[3]-48)*10+number[4]-48;
					if(year < 50)
					{
						year += 100;
					}
					switch(number[5]-48)
					{
					case 0:
						enroll = "體資生";
						break;
					case 1:
						enroll = "考試入學且系所不分班或甲班";
						break;
					case 2:
						enroll = "乙班";
						break;
					case 3:
						enroll = "丙班";
						break;
					case 4:
						enroll = "推甄入學";
						break;
					case 5:
						enroll = "僑生";
						break;
					case 6:
						enroll = "申請入學";
						break;
					case 7:
						enroll = "外籍生";
						break;
					case 8:
						enroll = "交換學生";
						break;
					case 9:
						enroll = "轉學生";
						break;
					}
					if(number[0] == 'H')
					{
						if(Score >= 0)
						{
							tmpTableData[flow-1] = new String[]{String.valueOf(flow++),String.valueOf(year),enroll,String.valueOf(Score)};
						}
						else
						{
							tmpTableData[flow-1] = new String[]{String.valueOf(flow++),String.valueOf(year),enroll,"空白"};
						}
					}
					workbook.close();
				}
			}else if(path.substring(path.length()-4).equals("xlsx") || path.substring(path.length()-4).equals("XLSX")){
				XSSFWorkbook workbook = new XSSFWorkbook(file);
				XSSFSheet sheet = workbook.getSheetAt(0);
				Iterator<Row> rowIterator = sheet.iterator();
				int flow = 1;
				findNum(rowIterator);
				rowIterator = sheet.iterator();
				Row row = rowIterator.next();
				tmpTableData = new String[num][4];
				while(rowIterator.hasNext())
				{
					row = rowIterator.next();
					byte[] number = row.getCell(StudentNumber).getStringCellValue().getBytes();
					double Score = -1;
					if(isScore(row))
					{
						Score = row.getCell(TotalScore).getNumericCellValue();
					}
					String enroll = "";
					int year = (number[3]-48)*10+number[4]-48;
					if(year < 50)
					{
						year += 100;
					}
					switch(number[5]-48)
					{
					case 0:
						enroll = "體資生";
						break;
					case 1:
						enroll = "考試入學且系所不分班或甲班";
						break;
					case 2:
						enroll = "乙班";
						break;
					case 3:
						enroll = "丙班";
						break;
					case 4:
						enroll = "推甄入學";
						break;
					case 5:
						enroll = "僑生";
						break;
					case 6:
						enroll = "申請入學";
						break;
					case 7:
						enroll = "外籍生";
						break;
					case 8:
						enroll = "交換學生";
						break;
					case 9:
						enroll = "轉學生";
						break;
					}
					System.out.println(flow+" "+new String(number)+" "+Score);
					if(number[0] == 'H')
					{
						if(Score >= 0)
						{
							tmpTableData[flow-1] = new String[]{String.valueOf(flow++),String.valueOf(year),enroll,String.valueOf(Score)};
						}
						else
						{
							tmpTableData[flow-1] = new String[]{String.valueOf(flow++),String.valueOf(year),enroll,"空白"};
						}
					}
					workbook.close();
				}
			}
			file.close();
		}catch(Exception e){}
	}
	boolean isScore(Row row)
	{
		boolean test = false;
		try
		{
			row.getCell(TotalScore).getNumericCellValue();
			test = true;
		}catch(Exception e){}
		return test;
	}
	void findNum(Iterator<Row> rowIterator)
	{
		try{
		Row row = rowIterator.next();
		while(rowIterator.hasNext())
		{
			row = rowIterator.next();
			byte[] number = row.getCell(StudentNumber).getStringCellValue().getBytes();
			if(number[0] == 'H')
			{
				num++;
			}
		}}catch(Exception e){}
	}
	void test2()
	{
		setBounds(100, 100, 450, 550);
		contentPane = new JPanel();
		setContentPane(contentPane);
		
		GridBagLayout gbl_contentPane = new GridBagLayout();
		gbl_contentPane.columnWidths = new int[]{0, 0, 0, 0, 0};
		gbl_contentPane.rowHeights = new int[]{0, 0, 0, 0, 0, 0};
		gbl_contentPane.columnWeights = new double[]{0.0, 1.0, 0.0, 0.0, Double.MIN_VALUE};
		gbl_contentPane.rowWeights = new double[]{0.0, 1.0, 0.0, 0.0, 0.0, Double.MIN_VALUE};
		contentPane.setLayout(gbl_contentPane);
		
		JLabel lblNewLabel = new JLabel("預覽輸出表格");
		GridBagConstraints gbc_lblNewLabel = new GridBagConstraints();
		gbc_lblNewLabel.anchor = GridBagConstraints.SOUTH;
		gbc_lblNewLabel.insets = new Insets(5, 0, 5, 0);
		gbc_lblNewLabel.gridwidth = 0;
		gbc_lblNewLabel.gridx = 1;
		gbc_lblNewLabel.gridy = 0;
		contentPane.add(lblNewLabel, gbc_lblNewLabel);
		
		JLabel label = new JLabel(" ");
		GridBagConstraints gbc_label = new GridBagConstraints();
		gbc_label.insets = new Insets(0, 0, 5, 5);
		gbc_label.gridx = 0;
		gbc_label.gridy = 1;
		contentPane.add(label, gbc_label);
		
		JTable table = new JTable(tmpTableData,book);
		GridBagConstraints gbc_table = new GridBagConstraints();
		gbc_table.insets = new Insets(0, 0, 5, 5);
		gbc_table.fill = GridBagConstraints.BOTH;
		gbc_table.gridwidth = 3;
		gbc_table.gridx = 1;
		gbc_table.gridy = 1;
		contentPane.add(new JScrollPane(table), gbc_table);
		
		JLabel label_1 = new JLabel(" ");
		GridBagConstraints gbc_label_1 = new GridBagConstraints();
		gbc_label_1.insets = new Insets(0, 0, 5, 0);
		gbc_label_1.gridx = 4;
		gbc_label_1.gridy = 1;
		contentPane.add(label_1, gbc_label_1);
		
		JLabel lblNewLabel_1 = new JLabel("輸出路徑");
		GridBagConstraints gbc_lblNewLabel_1 = new GridBagConstraints();
		gbc_lblNewLabel_1.insets = new Insets(0, 0, 7, 0);
		gbc_lblNewLabel_1.fill = GridBagConstraints.VERTICAL;
		gbc_lblNewLabel_1.gridwidth = 0;
		gbc_lblNewLabel_1.gridx = 1;
		gbc_lblNewLabel_1.gridy = 2;
		contentPane.add(lblNewLabel_1, gbc_lblNewLabel_1);
		
		textField = new JTextField();
		GridBagConstraints gbc_textField = new GridBagConstraints();
		gbc_textField.insets = new Insets(0, 0, 5, 5);
		gbc_textField.fill = GridBagConstraints.BOTH;
		gbc_textField.gridwidth = 2;
		gbc_textField.gridx = 1;
		gbc_textField.gridy = 3;
		contentPane.add(textField, gbc_textField);
		textField.setColumns(10);
		
		JButton btnNewButton_2 = new JButton("瀏覽");
		GridBagConstraints gbc_btnNewButton_2 = new GridBagConstraints();
		gbc_btnNewButton_2.fill = GridBagConstraints.HORIZONTAL;
		gbc_btnNewButton_2.insets = new Insets(0, 0, 5, 5);
		gbc_btnNewButton_2.gridx = 3;
		gbc_btnNewButton_2.gridy = 3;
		contentPane.add(btnNewButton_2, gbc_btnNewButton_2);
		btnNewButton_2.addMouseListener(new MouseAdapter() {
			public void mouseClicked(MouseEvent arg0) {
				try
				{
					JFileChooser chooser =new JFileChooser();
					chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
					chooser.setFileFilter(new MyFileFilter("請選擇輸出檔案的路徑"));
					int ret=chooser.showOpenDialog(null);
					if(ret==JFileChooser.APPROVE_OPTION)
					{
						textField.setText(chooser.getSelectedFile().getPath());
					}
				}catch(Exception e){
					System.out.println(e.toString());
				}
			}
		});
		
		JButton btnNewButton = new JButton("返回");
		GridBagConstraints gbc_btnNewButton = new GridBagConstraints();
		gbc_btnNewButton.anchor = GridBagConstraints.WEST;
		gbc_btnNewButton.insets = new Insets(0, 0, 5, 5);
		gbc_btnNewButton.gridx = 1;
		gbc_btnNewButton.gridy = 4;
		contentPane.add(btnNewButton, gbc_btnNewButton);
		btnNewButton.addMouseListener(new MouseAdapter() {
			public void mouseClicked(MouseEvent arg0) {
				try
				{
					contentPane.removeAll();
					contentPane.repaint();
					num = 0;
					test3();
				}catch(Exception e){
					System.out.println(e.toString());
				}
			}
		});
		JButton btnNewButton_1 = new JButton("完成");
		GridBagConstraints gbc_btnNewButton_1 = new GridBagConstraints();
		gbc_btnNewButton_1.insets = new Insets(0, 0, 5, 5);
		gbc_btnNewButton_1.fill = GridBagConstraints.BOTH;
		gbc_btnNewButton_1.gridwidth = 2;
		gbc_btnNewButton_1.gridx = 2;
		gbc_btnNewButton_1.gridy = 4;
		contentPane.add(btnNewButton_1, gbc_btnNewButton_1);
		btnNewButton_1.addMouseListener(new MouseAdapter() {
			public void mouseClicked(MouseEvent arg0) {
				try
				{
					XSSFWorkbook workbook = new XSSFWorkbook();
			        XSSFSheet sheet = workbook.createSheet("成績資料");
			        sheet.setColumnWidth(0,4000);
			        sheet.setColumnWidth(1,4000);
			        sheet.setColumnWidth(2,8000);
			        sheet.setColumnWidth(3,6000);
			        XSSFRow row = sheet.createRow(0);
			        XSSFCell cell = row.createCell(0);
			        cell.setCellValue(book[0]);
			        cell = row.createCell(1);
			        cell.setCellValue(book[1]);
			        cell = row.createCell(2);
			        cell.setCellValue(book[2]);
			        cell = row.createCell(3);
			        cell.setCellValue(book[3]);
			        for(int i = 1;i <= num;i++)
			        {
			        	row = sheet.createRow(i);
			        	for(int j = 0;j < 4;j++)
			        	{
			        		cell = row.createCell(j);
			        		if(j == 2)
			        		{
			        			
			        		}
			        		cell.setCellValue(tmpTableData[i-1][j]);
			        	}
			        }
			        FileOutputStream out = new FileOutputStream(textField.getText()+"成績輸出檔.xlsx");
			        workbook.write(out);
			        out.close();
			        WarningDialog f = new WarningDialog("完成");
			        System.exit(0);
				}catch(Exception e){
					e.printStackTrace();
					WarningDialog f = new WarningDialog(e.toString());
				}
			}
		});
	}
}