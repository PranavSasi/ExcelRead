package excelReadSample;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelSub {

	static FileInputStream a; //To Read Excel file
	static XSSFWorkbook b; //To Read WorkBook
	static XSSFSheet c;  //To Open Sheet
	
	public static String  readStringData(int i, int j) throws IOException {
		a=new FileInputStream("C:\\Users\\KuAp\\Desktop\\Java Excel Read Sample.xlsx");
		b=new XSSFWorkbook(a);
		c= b.getSheet("Sheet1");
		XSSFRow e= c.getRow(i);
		XSSFCell s= e.getCell(j);
		return s.getStringCellValue();
	}
	
	public static double readIntegerData(int i,int j) throws IOException {
		a=new FileInputStream("C:\\Users\\KuAp\\Desktop\\Java Excel Read Sample.xlsx");
		b=new XSSFWorkbook(a);
		c=b.getSheet("Sheet1");
		XSSFRow e= c.getRow(i);
		XSSFCell s= e.getCell(j);
		//int f=(int) s.getNumericCellValue();
		//return String.valueOf(f);
		return s.getNumericCellValue();
	}
	
}
