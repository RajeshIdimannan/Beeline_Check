package mainController;

import java.io.File;
import java.io.IOException;





import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;

import jxl.read.biff.BiffException;
import BeelineData.inputsheetManupulation;
import Beeline.BeelineReport;
import Beeline.DateProblem;
import BeelineData.inputsheetManupulation;
import configuration.SystemConfig;

public class MainController {
	
	public static String inputFile;
	public static String outputFile;
	public static void main(String[] args) throws ClassNotFoundException, IOException, InterruptedException, BiffException {
		//////////String of beeline manupulation//////////////
		//DateProblem date= new DateProblem();
		//int[] monYear= new int[]{};
//		String [] dateStr= new String[2];
		inputsheetManupulation inputSheetDate = new inputsheetManupulation();
//		////////commenting for the direct method for dates
//		try {
//			dateStr=inputSheetDate.gettingvaluedates();
//		} catch (BiffException e1) {
//			// TODO Auto-generated catch block
//			e1.printStackTrace();
//		}
//		//String [] strendDate=date.dateFinding(monYear[0],monYear[1]);
//		
//		for(int i=0;i<dateStr.length;i++){
//			System.out.println("Inside ecampact code"+ dateStr[i]);
//		}
//		
//		inputSheetDate.sheetManupulation(dateStr);
//		
		SystemConfig config= new SystemConfig();
		File dir = new File(".");
		//////declaring beeling project
		
		BeelineReport beeline=new BeelineReport();
		String[] mainArg=new String[2];
		//mainArg[0]= inputSheetDate.getFileName();
		try {
			inputFile = dir.getCanonicalPath() + "\\Input_Workbook.xls";
			outputFile = dir.getCanonicalPath() + "\\Test_Report.xls";
			config.configure();
			
		} catch (IOException e) {			
			e.printStackTrace();
		}
		///////changing code as beeline
		
		Thread.sleep(2000);
		
        String pathss=dir.getCanonicalPath()+"\\lib\\";
        System.out.println(pathss);
        
        boolean vbRun=false;
        String str =inputSheetDate.getFileName();
        System.out.println(str);
        try {
        	String[] vbParam = {"Wscript ",dir.getCanonicalPath()+"\\lib\\Excelexe.vbs", str};
            //Process p1 = Runtime.getRuntime().exec( "wscript " +dir.getCanonicalPath()+"\\lib\\Excelexe.vbs" );
        	Process p1 = Runtime.getRuntime().exec(vbParam);
        	p1.waitFor();
            vbRun=true;
         }
         catch( IOException e ) {
            System.out.println(e);
            System.exit(0);
         }
        if(vbRun==true){
        	//Thread.sleep(2000);
        	try {
				beeline.main(mainArg);
			} catch (Exception e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
        }
        DateFormat dateFormat = new SimpleDateFormat("dd-MM-yyyy");
		   //get current date time with Date()
		   Date date1 = new Date();
        
        try {
        	String filePath1 = System.getProperty("user.dir")+"\\BeelineReport@"+ dateFormat.format(date1);
        	String[] vbmsgParam = {"Wscript ",dir.getCanonicalPath()+"\\lib\\EmailSend.vbs ",filePath1};
        	Process p1 = Runtime.getRuntime().exec(vbmsgParam);
           //Process p1 = Runtime.getRuntime().exec( "wscript " +dir.getCanonicalPath()+"\\lib\\EmailSend.vbs",filePath1);        	
        	p1.waitFor();
            vbRun=true;
         }
         catch( IOException e ) {
            System.out.println(e);
            System.exit(0);
         }
	}
}
