package BeelineData;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Locale;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.format.Border;
import jxl.format.Colour;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import mainController.MainController;
import EndtoEnd.ExcelManipulation;

public class inputsheetManupulation {
	
	public void sheetManupulation(String[] arr){
	File dir = new File(".");
	Workbook workbook = null;
	WritableWorkbook writeSheet=null;
	WorkbookSettings wbSettings = new WorkbookSettings();
	wbSettings.setLocale(new Locale("en", "ER"));
	WritableSheet TestDataSheet = null;
	
	try{
		
		String inputFile = dir.getCanonicalPath() + "\\Input_Workbook.xls";
		System.out.println("inputFile");
		
		workbook =  Workbook.getWorkbook(new File(inputFile), wbSettings);
		writeSheet = Workbook.createWorkbook(new File(inputFile), workbook);
		TestDataSheet = writeSheet.getSheet("TEST DATA");
		System.out.println(arr[0]);
		System.out.println(arr[1]);
		Label startDate = new Label(4,1,arr[0]);
		Label endDate = new Label(6,1,arr[1]);
		TestDataSheet.addCell(startDate);
		TestDataSheet.addCell(endDate);
		writeSheet.write();
		writeSheet.close();
		workbook.close();
						
	}catch(Exception e){
		e.printStackTrace();
	}	
		
		
			
	}
	public int[] passingMonYear() throws BiffException, IOException{
		File dir = new File(".");
		Workbook workbook = null;
		
		WorkbookSettings wbSettings = new WorkbookSettings();
		wbSettings.setLocale(new Locale("en", "ER"));
		WritableSheet TestDataSheet = null;
		Sheet TestData=null;
		
		String inputFile = dir.getCanonicalPath() + "\\Input_Workbook.xls";
			workbook =  Workbook.getWorkbook(new File(inputFile));
		TestData = workbook.getSheet("TEST DATA");
		int[] monYear=new int[2];
		String labelValMon = TestData.getCell(9,1).getContents().toString();
		monYear[0] = Integer.parseInt(labelValMon);
		String labelValueMon = TestData.getCell(10,1).getContents().toString();
		monYear[1] = Integer.parseInt(labelValueMon);
		workbook.close();
		 return monYear;
	}
	public String getFileName() throws IOException, BiffException{
		File dir = new File(".");
		Workbook workbook = null;
		
		WorkbookSettings wbSettings = new WorkbookSettings();
		wbSettings.setLocale(new Locale("en", "ER"));
		WritableSheet TestDataSheet = null;
		Sheet TestData=null;
		
		String inputFile = dir.getCanonicalPath() + "\\Input_Workbook.xls";
			workbook =  Workbook.getWorkbook(new File(inputFile));
		TestData = workbook.getSheet("TEST DATA");
		String excelName;
		excelName = TestData.getCell(9,1).getContents().toString();
		return excelName;
		
	}
	public String[] gettingvaluedates() throws IOException, BiffException {
		File dir = new File(".");
		Workbook workbook = null;
		
		WorkbookSettings wbSettings = new WorkbookSettings();
		wbSettings.setLocale(new Locale("en", "ER"));
		WritableSheet TestDataSheet = null;
		Sheet TestData=null;
		
		String inputFile = dir.getCanonicalPath() + "\\Input_Workbook.xls";
			workbook =  Workbook.getWorkbook(new File(inputFile));
		TestData = workbook.getSheet("TEST DATA");
		String excelName;
		String dates[]=new String[2];
		 dates[0]=TestData.getCell(4,1).getContents().toString();
		 dates[1]=TestData.getCell(6,1).getContents().toString();
		SimpleDateFormat formatdate2 = new SimpleDateFormat("MM/dd/yyyy");
		formatdate2.format(dates[0]);
		formatdate2.format(dates[1]);
		
		return dates;
	}
}	
