package configuration;

import java.io.File;
import java.util.Locale;

import mainController.MainController;
import EndtoEnd.ExcelManipulation;
import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.WorkbookSettings;


public class SystemConfig {
	public static String ScreenshotsFlag,compatibilityFlag,systemFlag;
	public void configure(){
		
		Workbook workbook = null;
		WorkbookSettings wbSettings = new WorkbookSettings();
		wbSettings.setLocale(new Locale("en", "EN"));		
		try {
			workbook = Workbook.getWorkbook(new File(MainController.inputFile), wbSettings);
			Sheet CONFIG = workbook.getSheet("CONFIG");
					
			Cell screenshotsflag = CONFIG.findCell("SCREENSHOTS_FLAG");	
			Cell compatibilityflag = CONFIG.findCell("COMPATIBILITY_FLAG");
			Cell systemflag=CONFIG.findCell("SYSTEM");
			String Screenshots = CONFIG.getCell(screenshotsflag.getColumn()+1, screenshotsflag.getRow()).getContents();
			compatibilityFlag = CONFIG.getCell(compatibilityflag.getColumn()+1, compatibilityflag.getRow()).getContents();
			systemFlag =  CONFIG.getCell(systemflag.getColumn()+1, systemflag.getRow()).getContents();
			workbook.close();
			if(Screenshots.equalsIgnoreCase("All")){
				ScreenshotsFlag="Both";
				ExcelManipulation.E2ESystems();
			}else if(Screenshots.equalsIgnoreCase("Only Failed Steps")){
				ScreenshotsFlag="Fail";
				ExcelManipulation.E2ESystems();
			}else{
				ScreenshotsFlag="Off";
				ExcelManipulation.E2ESystems();
			}			
		}catch(Exception e){
			e.printStackTrace();
		}			
	}	
}
