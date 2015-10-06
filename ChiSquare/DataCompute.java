package v1;
import java.io.File;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.commons.math3.distribution.NormalDistribution;
import org.apache.commons.math3.geometry.euclidean.oned.Interval;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.commons.math3.stat.inference.ChiSquareTest;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
public class DataCompute {
	Cell cell;
	ArrayList<Double> Score = new ArrayList<Double>();
	double Average;
	double StandardDeviation;
	double[] expected;
	long[] observed;
	double GroupAmount = 11;
	ArrayList<Double> Separated = new ArrayList<Double>();
	ChiSquareTest ChiSquare = new ChiSquareTest();
	NormalDistribution Normal;
	double ChiSq;
	XSSFWorkbook newWorkbook = new XSSFWorkbook();
	XSSFWorkbook workbook;
	XSSFSheet sheet;
	String output = "";
	String file;
	double[] diff;
	double[] diffSqu;
	double totalDiffSqu = 0;
	double chiSquareValue = 0;
	double[] diffSquDivided;
	int printx = 4;
	int printy = 0;
	DataCompute(String inp,String outp,int groupnum){
		file = inp;
		output = outp;
		GroupAmount = groupnum;
		try{
			XSSFWorkbook workbook = new XSSFWorkbook(file);
			XSSFSheet sheet = workbook.getSheetAt(0);
			XSSFSheet sheet2 = newWorkbook.createSheet("卡方檢驗");
			Iterator<Row> rowIterator = sheet.iterator();
			Row row = rowIterator.next();
			int count = 0;
			XSSFRow row2 = sheet2.createRow(count++);
			for(int i = 0;i < 4;i++)
			{	
				XSSFCell cell2 = row2.createCell(i);
				cell2.setCellValue(row.getCell(i).getStringCellValue());
			}
			while(rowIterator.hasNext())
			{
				row = rowIterator.next();
				row2 = sheet2.createRow(count++);
				for(int i = 0;i < 4;i++)
				{	
					XSSFCell cell2 = row2.createCell(i);
					switch(row.getCell(i).getCellType())
					{
					case 0 :
						cell2.setCellValue(row.getCell(i).getNumericCellValue());
					      break;
					case 1 :
						cell2.setCellValue(row.getCell(i).getStringCellValue());
					      break;
					case 2 :
						cell2.setCellValue(row.getCell(i).getNumericCellValue());
					      break;
					case 3 :
					      break;
					case 4 :
						cell2.setCellValue(row.getCell(i).getBooleanCellValue());
					      break;
					case 5 :
					      break;
					}
				}
				double score = 0;
				cell = row.getCell(3);
				if(cell.getCellType() == 1 && cell.getStringCellValue().equals("空白"))
					cell.setCellValue(0.0);
				switch(cell.getCellType())
				{
				case 0 :
					score = cell.getNumericCellValue();
				      break;
				case 1 :
					score = Double.parseDouble(cell.getStringCellValue());
				      break;
				case 2 :
					score = cell.getNumericCellValue();
				      break;
				case 3 :
				      break;
				case 4 :
				      break;
				case 5 :
				      break;
				}
				Score.add(score);
			}
			System.out.println(Score.size());
			for(int i = 0;i < Score.size();i++)
			{
				System.out.print(Score.get(i)+" ");
			}System.out.println();
			ComputeAverage();
			ComputeStandardDeviation();
			setNormalDistribution();
			setSeparated();
			init_setExpected();
			initObserved();
			setObserved();
			compute();
			setNewWorkbook();
			int findsep = file.length()-5;
			for(;;)
			{
				findsep--;
				if(file.charAt(findsep) == '/' || file.charAt(findsep) == '\\')
				{
					findsep++;
					break;
				}
			}
			String name = file.substring(findsep, file.length()-5);
			FileOutputStream out = new FileOutputStream(output+"/"+name+".xlsx");
			newWorkbook.write(out);
			workbook.close();
		}catch(Exception e){e.printStackTrace();}
	}
	 void setNormalDistribution()
	{
		Normal = new NormalDistribution(Average,StandardDeviation);
	}
	 boolean isscore()
	{
		boolean test = false;
		try{
			cell.getNumericCellValue();
			test = true;
		}catch(Exception e){e.printStackTrace();}
		return test;
	}
	 void ComputeAverage()
	{
		double total = 0;
		for(int i = 0;i < Score.size();i++)
			total += Score.get(i);
		Average = total/Score.size();
	}
	 void ComputeStandardDeviation()
	{
		System.out.println(Score.size());
		double total = 0;
		for(int i = 0;i < Score.size();i++)
		{
			total += (Score.get(i)-Average) * (Score.get(i)-Average);
			System.out.println(Score.get(i));
		}
		total /= Score.size();
		StandardDeviation = Math.sqrt(total);
	}
	 void setSeparated(){
		for(double i = 1;i < GroupAmount;i++)
		{
			Separated.add(Normal.inverseCumulativeProbability(i/GroupAmount));
		}
		Separated.add(9999.0);
	}
	 void init_setExpected(){
		expected = new double[Separated.size()];
		for(int i = 0;i < Separated.size();i++)
		{
			expected[i] = (double)Score.size() / (double)GroupAmount;
		}
	}
	 void initObserved(){
		observed = new long[Separated.size()];
	}
	 void setObserved(){
		for(int i = 0;i < Score.size();i++)
		{
			double num = Score.get(i);
			boolean run = true;
			for(int j = 0;j < Separated.size() && run;j++)
			{
				if(num < Separated.get(j))
				{
					observed[j]++;
					run = false;
				}
			}
		}
	}
	 void compute(){
		 diff = new double[observed.length];
		 diffSqu = new double[observed.length];
		 diffSquDivided = new double[observed.length];
		 for(int i = 0;i < observed.length;i++)
		 {
			 diff[i] = Math.abs(observed[i] - expected[i]);
			 diffSqu[i] = diff[i] * diff[i];
			 diffSquDivided[i] = diffSqu[i] / expected[i];
			 totalDiffSqu += diffSqu[i];
			 chiSquareValue += diffSquDivided[i];
		 }
	 }
	 void setNewWorkbook(){
		XSSFSheet sheet = newWorkbook.getSheetAt(0);
		XSSFFont font = newWorkbook.createFont();
		font.setColor(IndexedColors.RED.getIndex());
		XSSFCellStyle style = newWorkbook.createCellStyle();
		style.setFont(font);
		sheet.setColumnWidth(0,4000);
        sheet.setColumnWidth(1,4000);
        sheet.setColumnWidth(2,8000);
        sheet.setColumnWidth(3,6000);
		for(int i = 3;i < 12;i++)
			sheet.setColumnWidth(i, 3000);
		sheet.setColumnWidth(12, 5000);
		XSSFRow row = sheet.getRow(printy);
		XSSFCell cell;
		String[] title = {"平均數","標準差","組編號","組距","觀察值","期望值","差值","差值平方","差值平方/期望值"};
		cell = row.createCell(printx);
		cell.setCellValue("P值");
		cell = row.createCell(printx+1);
		System.out.println(expected.length + " " + observed.length);
		cell.setCellValue(ChiSquare.chiSquareTest(expected, observed));
		row = sheet.getRow(printy+1);
		cell = row.createCell(printx);
		cell.setCellValue("卡方值");
		cell = row.createCell(printx+1);
		cell.setCellValue(chiSquareValue);
		row = sheet.getRow(printy+2);
		cell = row.createCell(printx);
		cell.setCellValue("自由度");
		cell = row.createCell(printx+1);
		cell.setCellValue(observed.length-1);
		row = sheet.getRow(printy+5);
		for(int i = 0;i < 9;i++)
		{
			cell = row.createCell(printx+i);
			cell.setCellValue(title[i]);
			cell.setCellStyle(style);
		}
		row = sheet.getRow(printy+6);
		cell = row.createCell(printx);
		cell.setCellValue(Average);
		cell.setCellStyle(style);
		cell = row.createCell(printx+1);
		cell.setCellValue(StandardDeviation);
		cell.setCellStyle(style);
		cell = row.createCell(printx+2);
		cell.setCellValue(1);
		cell.setCellStyle(style);
		cell = row.createCell(printx+3);
		cell.setCellValue(Separated.get(0));
		cell.setCellStyle(style);
		cell = row.createCell(printx+4);
		cell.setCellValue(observed[0]);
		cell.setCellStyle(style);
		cell = row.createCell(printx+5);
		cell.setCellValue(expected[0]);
		cell.setCellStyle(style);
		cell = row.createCell(printx+6);
		cell.setCellValue(diff[0]);
		cell.setCellStyle(style);
		cell = row.createCell(printx+7);
		cell.setCellValue(diffSqu[0]);
		cell.setCellStyle(style);
		cell = row.createCell(printx+8);
		cell.setCellValue(diffSquDivided[0]);
		cell.setCellStyle(style);
		for(int i = 1;i < observed.length;i++)
		{
			row = sheet.getRow(printy+6+i);
			cell = row.createCell(printx+2);
			cell.setCellValue(i+1);
			cell.setCellStyle(style);
			cell = row.createCell(printx+3);
			cell.setCellValue(Separated.get(i));
			cell.setCellStyle(style);
			cell = row.createCell(printx+4);
			cell.setCellValue(observed[i]);
			cell.setCellStyle(style);
			cell = row.createCell(printx+5);
			cell.setCellValue(expected[i]);
			cell.setCellStyle(style);
			cell = row.createCell(printx+6);
			cell.setCellValue(diff[i]);
			cell.setCellStyle(style);
			cell = row.createCell(printx+7);
			cell.setCellValue(diffSqu[i]);
			cell.setCellStyle(style);
			cell = row.createCell(printx+8);
			cell.setCellValue(diffSquDivided[i]);
			cell.setCellStyle(style);
		}
	}
}
