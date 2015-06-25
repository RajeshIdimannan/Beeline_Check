package EndtoEnd;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

import mainController.MainController;

import org.apache.commons.lang3.StringUtils;
import org.openqa.selenium.WebDriver;

import configuration.SystemConfig;
import jxl.Cell;
import jxl.Sheet;
import jxl.SheetSettings;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.format.Alignment;
import jxl.format.Border;
import jxl.format.BorderLineStyle;
import jxl.format.Colour;
import jxl.format.UnderlineStyle;
import jxl.format.VerticalAlignment;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

public class ExcelManipulation {
	
	private  String inputFile;
	private  String outputFile;
	
	
	public String getInputFile() {
		return inputFile;
	}

	public void setInputFile(String inputFile) {
		inputFile = this.inputFile;
	}

	public String getOutputFile() {
		return outputFile;
	}

	public void setOutputFile(String outputFile) {
		outputFile = this.outputFile;
	}

	public static ArrayList<String> getTargetList(String content) {
		ArrayList<String> list = new ArrayList<String>();
		String[] a = content.split("\n");
		for (int i = 0; i < a.length; i++) {
			list.add(a[i]);
		}
		return list;
	}
	
	public void copyExcel()
			throws IOException {
		Workbook workbook = null;
		WritableWorkbook copy;
		WorkbookSettings wbSettings = new WorkbookSettings();
		wbSettings.setLocale(new Locale("en", "EN"));
		try {
			workbook = Workbook.getWorkbook(new File(MainController.inputFile), wbSettings);			
			copy = Workbook.createWorkbook(new File(MainController.outputFile));
			
			///////////Creating DRIVER Sheet///////////////////
			copy.createSheet("DRIVER", copy.getNumberOfSheets());
			WritableSheet DriverSheet = copy.getSheet("DRIVER");
			SheetSettings dst=	DriverSheet.getSettings();		
			dst.setShowGridLines(false);
			Label SNO= new Label(0, 0,"S NO" ,
					createFormattedCell(11, WritableFont.createFont("calibri"), true, false, null, Border.ALL,
							null, null, Colour.WHITE, Colour.DARK_BLUE));
			Label track= new Label(1, 0,"TRACK/MODULE" ,
					createFormattedCell(11, WritableFont.createFont("calibri"), true, false, null, Border.ALL,
							null, null, Colour.WHITE, Colour.DARK_BLUE));
			Label ScenarioName= new Label(2, 0,"SCENARIO NAME" ,
					createFormattedCell(11, WritableFont.createFont("calibri"), true, false, null, Border.ALL,
							null, null, Colour.WHITE, Colour.DARK_BLUE));
			DriverSheet.addCell(SNO);
			DriverSheet.addCell(track);
			DriverSheet.addCell(ScenarioName);
			
			DriverSheet.setColumnView(0, 10);
			DriverSheet.setColumnView(1, 28);
			DriverSheet.setColumnView(2, 38);
			DriverSheet.setColumnView(3, 38);
			DriverSheet.setColumnView(4, 38);
			DriverSheet.setColumnView(5, 38);
			DriverSheet.setColumnView(6, 38);
			
			/////////// Creating and Writing Result Sheet Labels///////////////
			copy.createSheet("OUTPUT_VALUES", copy.getNumberOfSheets());
			WritableSheet ResultSheet = copy.getSheet("OUTPUT_VALUES");
			SheetSettings st=	ResultSheet.getSettings();		
			st.setShowGridLines(false);
			
			Label RSNO= new Label(0, 0,"SNO" ,
					createFormattedCell(10,  WritableFont.createFont("calibri"), true, false, null,  Border.ALL,
							null, null, Colour.GRAY_80, Colour.VERY_LIGHT_YELLOW));
			Label Scenario= new Label(1, 0,"SCENARIO" ,
					createFormattedCell(10,  WritableFont.createFont("calibri"), true, false, null,  Border.ALL,
							null, null, Colour.GRAY_80, Colour.VERY_LIGHT_YELLOW));
			
			ResultSheet.addCell(RSNO);
			ResultSheet.addCell(Scenario);
			
			ResultSheet.setColumnView(0, 20);
			ResultSheet.setColumnView(1, 20);
								
			//////////// Test case wise Report//////////////////////
			copy.createSheet("TESTCASEWISE_REPORT", copy.getNumberOfSheets());
			WritableSheet ScenarioSheet = copy.getSheet("TESTCASEWISE_REPORT");
			SheetSettings sst=	ScenarioSheet.getSettings();		
			sst.setShowGridLines(false);			
			
			Label ScenarioName1= new Label(0, 0,"SCENARIO NAME" ,
					createFormattedCell(10,  WritableFont.createFont("calibri"), true, false, null,  Border.ALL,
							null, null, Colour.GRAY_80, Colour.LIGHT_TURQUOISE2));
			Label fnname= new Label(1, 0,"SYSTEM/FUNCTION NAME" ,
					createFormattedCell(10,  WritableFont.createFont("calibri"), true, false, null,  Border.ALL,
							null, null, Colour.GRAY_80, Colour.LIGHT_TURQUOISE2));
			Label status= new Label(2, 0,"EXECUTION STATUS" ,
					createFormattedCell(10,  WritableFont.createFont("calibri"), true, false, null,  Border.ALL,
							null, null, Colour.GRAY_80, Colour.LIGHT_TURQUOISE2));
			Label time= new Label(3, 0,"EXECUTION TIME (in sec's)" ,
					createFormattedCell(10,  WritableFont.createFont("calibri"), true, false, null,  Border.ALL,
							null, null, Colour.GRAY_80, Colour.LIGHT_TURQUOISE2));
			
			ScenarioSheet.addCell(ScenarioName1);
			ScenarioSheet.addCell(fnname);
			ScenarioSheet.addCell(status);
			ScenarioSheet.addCell(time);
			
			ScenarioSheet.setColumnView(0, 48);
			ScenarioSheet.setColumnView(1, 18);
			ScenarioSheet.setColumnView(2, 22);
			ScenarioSheet.setColumnView(3, 14);	
			ScenarioSheet.setColumnView(4, 22);
			ScenarioSheet.setColumnView(5, 14);	
			ScenarioSheet.setColumnView(6, 22);
			ScenarioSheet.setColumnView(7, 14);	
			
        	/////////// Creating and Writing Detailed Report Sheet Labels/////////
			copy.createSheet("DETAILED REPORT", copy.getNumberOfSheets());
			WritableSheet ReportSheet = copy.getSheet("DETAILED REPORT");
			SheetSettings st1=	ReportSheet.getSettings();		
			st1.setShowGridLines(false);
			
			Label ScenarioReport= new Label(0, 0,"SCENARIO NAME" ,
					createFormattedCell(10,  WritableFont.createFont("calibri"), true, false, null,  Border.ALL,
							null, null, Colour.GRAY_80, Colour.LIGHT_TURQUOISE2));
			Label functionname= new Label(1, 0,"FUNCTION NAME" ,
					createFormattedCell(10,  WritableFont.createFont("calibri"), true, false, null,  Border.ALL,
							null, null, Colour.GRAY_80, Colour.LIGHT_TURQUOISE2));
			Label stepname= new Label(2, 0,"STEP DESCRIPTION" ,
					createFormattedCell(10,  WritableFont.createFont("calibri"), true, false, null,  Border.ALL,
							null, null, Colour.GRAY_80, Colour.LIGHT_TURQUOISE2));
			Label result= new Label(3, 0,"RESULT" ,
					createFormattedCell(10,  WritableFont.createFont("calibri"), true, false, null,  Border.ALL,
							null, null, Colour.GRAY_80, Colour.LIGHT_TURQUOISE2));
			Label errlog= new Label(4, 0,"ERROR LOG" ,
					createFormattedCell(10,  WritableFont.createFont("calibri"), true, false, null,  Border.ALL,
							null, null, Colour.GRAY_80, Colour.LIGHT_TURQUOISE2));	
			Label Screenshots= new Label(5, 0,"SCREENSHOTS" ,
					createFormattedCell(10,  WritableFont.createFont("calibri"), true, false, null,  Border.ALL,
							null, null, Colour.GRAY_80, Colour.LIGHT_TURQUOISE2));
			
			ReportSheet.addCell(ScenarioReport);
			ReportSheet.addCell(functionname);
			ReportSheet.addCell(stepname);			
			ReportSheet.addCell(result);
			ReportSheet.addCell(errlog);
			ReportSheet.addCell(Screenshots);
			ReportSheet.setColumnView(0, 60);
			ReportSheet.setColumnView(1, 20);
			ReportSheet.setColumnView(2, 20);
			ReportSheet.setColumnView(3, 20);
			ReportSheet.setColumnView(4, 40);
			ReportSheet.setColumnView(5, 45);
			ReportSheet.setColumnView(6, 20);
			ReportSheet.setColumnView(7, 40);
			ReportSheet.setColumnView(8, 45);
			ReportSheet.setColumnView(9, 20);
			ReportSheet.setColumnView(10, 40);
			ReportSheet.setColumnView(11, 45);
			
			///////////////////////////////////////////PIXEL REPORT////////////////////////////////////////////
			if(SystemConfig.compatibilityFlag.equalsIgnoreCase("On")){
				copy.createSheet("PIXEL_REPORT", copy.getNumberOfSheets());
				WritableSheet pixelSheet = copy.getSheet("PIXEL_REPORT");
				SheetSettings pxlsettings=	pixelSheet.getSettings();		
				pxlsettings.setShowGridLines(false);
				Label Scenariolabel= new Label(0, 0,"SCENARIO NAME" ,
						createFormattedCell(10,  WritableFont.createFont("calibri"), true, false, null,  Border.ALL,
								null, null, Colour.GRAY_80, Colour.TAN));
				Label objlabel= new Label(1, 0,"OBJECT NAME" ,
						createFormattedCell(10,  WritableFont.createFont("calibri"), true, false, null,  Border.ALL,
								null, null, Colour.GRAY_80, Colour.TAN));
				Label fontFamily= new Label(2, 0,"FONT FAMILY" ,
						createFormattedCell(10,  WritableFont.createFont("calibri"), true, false, null,  Border.ALL,
								null, null, Colour.GRAY_80, Colour.TAN));
				Label fontColor= new Label(3, 0,"FONT COLOR" ,
						createFormattedCell(10,  WritableFont.createFont("calibri"), true, false, null,  Border.ALL,
								null, null, Colour.GRAY_80, Colour.TAN));
				Label fontStyle= new Label(4, 0,"FONT STYLE" ,
						createFormattedCell(10,  WritableFont.createFont("calibri"), true, false, null,  Border.ALL,
								null, null, Colour.GRAY_80, Colour.TAN));
				Label fontWeight= new Label(5, 0,"FONT WEIGHT" ,
						createFormattedCell(10,  WritableFont.createFont("calibri"), true, false, null,  Border.ALL,
								null, null, Colour.GRAY_80, Colour.TAN));
				Label fontSize= new Label(6, 0,"FONT SIZE" ,
						createFormattedCell(10,  WritableFont.createFont("calibri"), true, false, null,  Border.ALL,
								null, null, Colour.GRAY_80, Colour.TAN));
				Label letTop= new Label(7, 0,"LEFT TOP" ,
						createFormattedCell(10,  WritableFont.createFont("calibri"), true, false, null,  Border.ALL,
								null, null, Colour.GRAY_80, Colour.TAN));
				Label rightBottom= new Label(8, 0,"RIGHT BOTTOM" ,
						createFormattedCell(10,  WritableFont.createFont("calibri"), true, false, null,  Border.ALL,
								null, null, Colour.GRAY_80, Colour.TAN));
				
				pixelSheet.addCell(Scenariolabel);
				pixelSheet.addCell(objlabel);
				pixelSheet.addCell(fontFamily);
				pixelSheet.addCell(fontColor);
				pixelSheet.addCell(fontStyle);
				pixelSheet.addCell(fontWeight);
				pixelSheet.addCell(fontSize);
				pixelSheet.addCell(letTop);
				pixelSheet.addCell(rightBottom);
			}
			copy.write();			
			copy.close();
			workbook.close();
		} catch (Exception e) {
			System.out.println(e.getMessage());
		}
	}
	
	public static WritableCellFormat createFormattedCell(int pointSize,
			WritableFont.FontName fontName, boolean isBold, boolean italic,
			UnderlineStyle underLineStyle, Border border,
			BorderLineStyle lineStyle, Alignment alignment, Colour fontcolor,
			Colour background) throws WriteException {
		WritableFont font = new WritableFont(fontName != null ? fontName
				: WritableFont.ARIAL, pointSize, isBold ? WritableFont.BOLD
				: WritableFont.NO_BOLD, italic,
				underLineStyle != null ? underLineStyle
						: UnderlineStyle.NO_UNDERLINE);
		font.setColour(fontcolor);

		WritableCellFormat writableCellFormat = new WritableCellFormat(font);
		if (lineStyle == null) {
			lineStyle = BorderLineStyle.THIN;
		}
		if (border == null) {
			border = Border.NONE;
		}
		if (alignment == null) {
			alignment = Alignment.CENTRE;
		}
		writableCellFormat.setWrap(true);
		writableCellFormat.setBorder(border, lineStyle, Colour.BLACK);
		writableCellFormat.setAlignment(alignment);
		writableCellFormat.setBackground(background);
		writableCellFormat.setVerticalAlignment(VerticalAlignment.CENTRE);
		return writableCellFormat;
	}
	
	
	public static E2EDriver readtrack(E2EDriver e2eDriver,String systemName){
		try{			
			Workbook workbook = null;			
			Sheet systemSheet = null;										
			workbook = Workbook.getWorkbook(new File(MainController.inputFile));						
			Sheet CONFIG = workbook.getSheet("CONFIG");						
			systemSheet=workbook.getSheet(systemName);
			int ecount=0;
			ArrayList<eStoreData> estorelist = new ArrayList<eStoreData>();
			if(e2eDriver.geteStoreData()!=null){
				ecount=e2eDriver.geteStoreData().size();
				estorelist=e2eDriver.geteStoreData();
			}
			Cell scenarioncell = systemSheet.findCell(e2eDriver.getScenario());
			for(int scenarioiterator = scenarioncell.getRow(); scenarioncell!=null && scenarioiterator<systemSheet.getRows() ;scenarioiterator++){
				if(!(systemSheet.getCell(0, scenarioiterator).getContents().equalsIgnoreCase(e2eDriver.getScenario().toString()))){
					break;
				}else{
					
					eStoreData estorefn = new eStoreData();					
					estorefn.setSystemName(systemName);
					estorefn.setFnname(systemSheet.getCell(2, scenarioiterator).getContents().toString());
					String content=systemSheet.getCell(3, scenarioiterator).getContents().toString();
					estorefn.setParameter(content);
					ArrayList<String> list = new ArrayList<String>();
					if(content.contains(",") && !content.isEmpty()){
						String[] a = content.split(",");
						for (int index = 0; index < a.length; index++) {
							list.add(a[index]);
						}
					}else if(!content.contains(",")){
						list.add(content);
					}
					
					
					Sheet fnlib=workbook.getSheet("FN_LIB");
					Sheet OR=workbook.getSheet("OR");
					int fncount=0;
					ArrayList<SeleneseData> seleneselist=new ArrayList<SeleneseData>();
					Cell functionCell=fnlib.findCell(estorefn.getFnname().toString());										
						for(int rowIterator = functionCell.getRow();functionCell != null && rowIterator<fnlib.getRows() ;rowIterator++){
							functionCell=fnlib.getCell(0, rowIterator);
							if(!(fnlib.getCell(0, rowIterator).getContents().equalsIgnoreCase(estorefn.getFnname().toString()))){
								break;
							}else{
								
								SeleneseData seleneseData = new SeleneseData();
								seleneseData.setStep(fnlib.getCell(1, rowIterator).getContents().toString());
								seleneseData.setErrorFlag(fnlib.getCell(2, rowIterator).getContents());
								if (fnlib.getCell(3, rowIterator).getContents().toString().isEmpty()) {
									seleneseData.setTargetList(getTargetList(fnlib.getCell(3, rowIterator).getContents().toString()));
									seleneseData.setTarget((String) seleneseData.getTargetList().get(0));

								} else {
									String Objects = fnlib.getCell(3, rowIterator).getContents().toString();
									if(Objects.contains(",") && !Objects.isEmpty()){
										String[] obj = Objects.split(",");
										ArrayList<String> targetList = new ArrayList<String>();
										for (int index = 0; index < obj.length; index++) {											
											Cell objCell = OR.findCell(obj[index]);
											if (objCell == null) {
												seleneseData.setTarget("");
											} else {
												String realObj= OR.getCell(2,objCell.getRow()).getContents().toString();
												targetList.add(realObj);																								
											}
										}
										seleneseData.setTargetList(targetList);
									}else if(!Objects.contains(",")){
										Cell objCell = OR.findCell(Objects);
										if (objCell == null) {
											seleneseData.setTarget("Object not found in object repository");
										} else {
											seleneseData.setTargetList(getTargetList(OR.getCell(2,objCell.getRow()).getContents().toString()));
											seleneseData.setTarget((String) seleneseData.getTargetList().get(0));
										}
									}
								}
								seleneseData.setCommand(fnlib.getCell(4, rowIterator).getContents().toString());
								String parameterName = fnlib.getCell(5, rowIterator).getContents().toString();
																														
								if(parameterName.startsWith("param") && !content.isEmpty()){
									parameterName = list.get(0);
									list.remove(0);
									seleneseData.setValue(parameterName);									
								}else{
									seleneseData.setValue(parameterName);
								}	
								
								 if(!parameterName.isEmpty() && !seleneseData.getCommand().equalsIgnoreCase("storevalue")){
										if(!StringUtils.isNumeric(parameterName)){
											Sheet TestData = workbook.getSheet("TEST DATA");
											if(TestData.findCell(parameterName) != null){
												if(!(TestData.getCell(TestData.findCell(parameterName).getColumn(), TestData.findCell(e2eDriver.getScenario()).getRow()).getContents().toString()).isEmpty()){
													seleneseData.setValue((TestData.getCell(TestData.findCell(parameterName).getColumn(), TestData.findCell(e2eDriver.getScenario()).getRow()).getContents().toString()));
												}else{
													seleneseData.setValue(parameterName);
												}												
											}else if (CONFIG.findCell(parameterName) != null){
												seleneseData.setValue(CONFIG.getCell(CONFIG.findCell(parameterName).getColumn()+1,CONFIG.findCell(parameterName).getRow()).getContents().toString());
											}else{
												seleneseData.setValue(parameterName);
											}														
										}
									 }
								 
								seleneselist.add(fncount, seleneseData);
								fncount++;
								estorefn.setSeleneseData(seleneselist);
							}											
						}						
					estorelist.add(ecount, estorefn);										
					ecount++;					
					e2eDriver.seteStoreData(estorelist);					
				}
			}
			workbook.close();
		}catch(Exception e){
			e.printStackTrace();
		}
		
		return e2eDriver;
	}
	
	
	public static ArrayList<E2EDriver> readExcel() {
		Workbook workbook = null;
		ArrayList<E2EDriver> listObjExcelData = new ArrayList<E2EDriver>();
		String systemName=null;
		
		try {
			workbook = Workbook.getWorkbook(new File(MainController.inputFile));			
			Sheet DRIVER = workbook.getSheet(0);
			Sheet CONFIG = workbook.getSheet("CONFIG");
			int count =0;
			for(int i=1;i<DRIVER.getRows();i++){				
				try {	
					if( DRIVER.getCell(3, i).getContents().toString().equalsIgnoreCase("Yes")){
						E2EDriver e2eDriver = new E2EDriver();
						List<String> URL = new ArrayList<>();
						List<String> browser = new ArrayList<>();
						List<String> sysname = new ArrayList<>();
						e2eDriver.setTrack(DRIVER.getCell(1, i).getContents().toString());
						e2eDriver.setScenario(DRIVER.getCell(2, i).getContents().toString());						
						Sheet Track=workbook.getSheet(e2eDriver.getTrack());
						Cell scenarioName=Track.findCell(e2eDriver.getScenario());						
						Cell systemURL,systemBrowser;
						////Single System/////
						if(SystemConfig.systemFlag.equalsIgnoreCase("Specific Track")){
							systemURL=CONFIG.findCell("URL");
							URL.add(CONFIG.getCell(systemURL.getColumn()+1,systemURL.getRow()).getContents().toString());
							e2eDriver.setBaseurl(URL);
							systemBrowser=CONFIG.findCell("BROWSER");
							browser.add(CONFIG.getCell(systemBrowser.getColumn()+1, systemBrowser.getRow()).getContents());
							e2eDriver.setBrowser(browser);
							systemName=Track.getName();
							e2eDriver=	readtrack(e2eDriver,systemName);
							e2eDriver.setSystemCount(1);
							sysname.add(systemName);
							e2eDriver.setSystemName(sysname);
						}
						/////Multiple Systems/////
						else if (SystemConfig.systemFlag.equalsIgnoreCase("E2E Systems")){							
							for(int j=1;((j<=Track.getColumns()-2));j++){
								if((Track.getCell(scenarioName.getColumn()+j, scenarioName.getRow()).getContents().equalsIgnoreCase("Yes"))){									
									systemName=Track.getCell(scenarioName.getColumn()+j, 0).getContents().toString();
									systemURL=CONFIG.findCell(systemName+"_URL");
									URL.add(CONFIG.getCell(systemURL.getColumn()+1,systemURL.getRow()).getContents().toString());
									systemBrowser=CONFIG.findCell(systemName+"_BROWSER");
									browser.add(CONFIG.getCell(systemBrowser.getColumn()+1, systemBrowser.getRow()).getContents());
									e2eDriver=	readtrack(e2eDriver,systemName);
									sysname.add(systemName);
								}
							}
							e2eDriver.setBaseurl(URL);
							e2eDriver.setBrowser(browser);
							e2eDriver.setSystemCount(browser.size());
							e2eDriver.setSystemName(sysname);
						}
						listObjExcelData.add(count,e2eDriver);
						count++;
					}
				}catch(Exception e){
					e.printStackTrace();
				}				
			}
			workbook.close();			
		} catch (Exception e) {
			e.printStackTrace();
		}
		return listObjExcelData;
	}
	
	
	public static void E2ESystems(){
		try {
			WebDriver driver = null;
			boolean fnconditionflag=false;
			boolean errorFlag=false;
			
			 /////////////////////////////////////// Data Retrieval From DB ///////////////////////////////////
			/*Process p = Runtime.getRuntime().exec("cscript c:\\test.vbs");	
			p.waitFor();
			BufferedReader input =new BufferedReader(new InputStreamReader(p.getInputStream()));
		        String line;
		       
				while ((line = input.readLine()) != null){
					l++;
					if(l>3) 
						orderno=line;
				}
			input.close();	
			inputFile="C:\\Documents and Settings\\A530883\\Desktop\\E2E\\E2E_Input.xls";
			outputFile="C:\\Documents and Settings\\A530883\\Desktop\\E2E\\output.xls";*/
			
			File dir = new File(".");			
			DriverExecution exeExcel = new DriverExecution();
			ArrayList<E2EDriver> ElementArray = new ArrayList<E2EDriver>();
			Timer time=new Timer();
			time.start();
			ElementArray=readExcel();
			time.end();
			System.out.println(time.getTotalTime());
			
			
			ExcelManipulation rdExcel=new ExcelManipulation();
			rdExcel.copyExcel();
			
			for(E2EDriver edata : ElementArray){													//Scenario Loop
				for(int syscount=1;syscount<=edata.getSystemCount();syscount++){						//System Loop	
					ArrayList<eStoreData> estorelist =edata.geteStoreData();
					List<String> browser =edata.getBrowser();
					List<String> URL=edata.getBaseurl();
					String[] browserArray = browser.get(syscount-1).split(",");
					for(int browsercount=0;browsercount<browserArray.length;browsercount++){			//Browser Loop
						String browserVersion = exeExcel.initDriver(browserArray[browsercount],"");
						driver=exeExcel.loadApplication(URL.get(syscount-1));
						Timer scenarioTimer = new Timer();
						Timer functionTimer = new Timer();
						scenarioTimer.start();
						String systemName = null;
						int testcasewisecount=1,stepcount=1;
						rdExcel.writeTestcasewiseReportBrowserLabel(browsercount,edata.getScenario(),edata.getSystemName().get(syscount-1),browserVersion);
						rdExcel.writeDetailReportBrowserLabel(browsercount,edata.getScenario(),edata.getSystemName().get(syscount-1),browserVersion);
						for(eStoreData esdata:estorelist){												//Function Loop
							if(esdata.getSystemName().equalsIgnoreCase(edata.getSystemName().get(syscount-1))){
								functionTimer.start();
								ArrayList<SeleneseData> seleneselist=esdata.getSeleneseData();								
								seleneselist = exeExcel.runserver(seleneselist,dir.getCanonicalPath(),browserVersion,edata.getScenario());
								functionTimer.end();
								systemName=esdata.getSystemName();								
								fnconditionflag=rdExcel.writeDetailedResults(browsercount,stepcount,seleneselist,edata,esdata.getFnname());
								stepcount=stepcount+esdata.getSeleneseData().size();
								
								fnconditionflag=rdExcel.writeTestCaseWiseResults(browsercount,testcasewisecount,fnconditionflag,edata,functionTimer.getTotalTime(),esdata.getFnname(),systemName);								
								testcasewisecount++;
								for (int i = 0; seleneselist!= null && i < seleneselist.size(); i++){
									SeleneseData seleneseData = (SeleneseData) seleneselist.get(i);							
									if(!seleneseData.isOverallStatus()){
										errorFlag=true;
									}							
								}
								if(errorFlag){
									break;
								}
							}								
						}					
						scenarioTimer.end();
						rdExcel.writeDriverResult(fnconditionflag,scenarioTimer.getTotalTime(),edata,systemName);
						driver.quit();
					}
				}
			}			
		}catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	public void writeTestcasewiseReportBrowserLabel(int browsercount,String scenarioName,String systemName,String browserVersion){
		Workbook workbook = null;
		WritableWorkbook copy = null;
		WorkbookSettings wbSettings = new WorkbookSettings();
		wbSettings.setLocale(new Locale("en", "ER"));
		Label Scenario,system,browserversion = null;//,result = null,errorLog,screenshots;
		int j=browsercount;
		try{
			workbook = Workbook.getWorkbook(new File(MainController.outputFile), wbSettings);
			copy = Workbook.createWorkbook(new File(MainController.outputFile), workbook);
			
			/////////////////////// Detailed Report /////////////////////////
			WritableSheet reportSheet = copy.getSheet("TESTCASEWISE_REPORT");
			
			if(browsercount==0){
				
				Scenario= new Label(0, reportSheet.getRows(),scenarioName ,
						createFormattedCell(10,  WritableFont.createFont("calibri"), true, false, null,  Border.ALL,
								null,  Alignment.LEFT, Colour.GRAY_80, Colour.LIGHT_TURQUOISE2));											
				system = new Label(1, reportSheet.getRows(),systemName.toUpperCase() ,
						createFormattedCell(10,  WritableFont.createFont("calibri"), true, false, null,  Border.ALL,
								null, null, Colour.WHITE, Colour.GRAY_50));
				reportSheet.mergeCells(2,reportSheet.getRows(), 3,reportSheet.getRows());
				browserversion = new Label(2, reportSheet.getRows()-1,browserVersion.toUpperCase() ,
						createFormattedCell(10,  WritableFont.createFont("calibri"), true, false, null,  Border.ALL,
								null, null, Colour.WHITE, Colour.GRAY_50));
				reportSheet.addCell(Scenario);
				reportSheet.addCell(system);
			}else{
				if(browsercount==2){
					j+=1;
				}
				Label status= new Label(j+3, 0,"EXECUTION STATUS" ,
						createFormattedCell(10,  WritableFont.createFont("calibri"), true, false, null,  Border.ALL,
								null, null, Colour.GRAY_80, Colour.LIGHT_TURQUOISE2));
				Label time= new Label(j+4, 0,"EXECUTION TIME (in sec's)" ,
						createFormattedCell(10,  WritableFont.createFont("calibri"), true, false, null,  Border.ALL,
								null, null, Colour.GRAY_80, Colour.LIGHT_TURQUOISE2));
				Cell scenariocell = reportSheet.findCell(scenarioName);
				reportSheet.mergeCells(j+3,scenariocell.getRow(), j+4,scenariocell.getRow());
				browserversion = new Label(j+3,scenariocell.getRow(),browserVersion.toUpperCase() ,
						createFormattedCell(10,  WritableFont.createFont("calibri"), true, false, null,  Border.ALL,
								null, null, Colour.WHITE, Colour.GRAY_50));
				reportSheet.addCell(status);
				reportSheet.addCell(time);
			}
			reportSheet.addCell(browserversion);
			
			copy.write();
			copy.close();
			workbook.close();
		}catch(Exception e){
			e.printStackTrace();
		}
	}
	
	public void writeDetailReportBrowserLabel(int browsercount,String scenarioName,String systemName,String browserVersion){
		Workbook workbook = null;
		WritableWorkbook copy = null;
		WorkbookSettings wbSettings = new WorkbookSettings();
		wbSettings.setLocale(new Locale("en", "ER"));
		Label Scenario,system,browserversion = null;//,result = null,errorLog,screenshots;
		int j=browsercount;
		try{
			workbook = Workbook.getWorkbook(new File(MainController.outputFile), wbSettings);
			copy = Workbook.createWorkbook(new File(MainController.outputFile), workbook);
			
			/////////////////////// Detailed Report /////////////////////////
			WritableSheet reportSheet = copy.getSheet("DETAILED REPORT");
			if(browsercount==0){
			
				Scenario= new Label(0, reportSheet.getRows(),scenarioName ,
						createFormattedCell(10,  WritableFont.createFont("calibri"), true, false, null,  Border.ALL,
								null,  Alignment.LEFT, Colour.GRAY_80, Colour.LIGHT_TURQUOISE2));
				
				reportSheet.mergeCells(1, reportSheet.getRows(), 2, reportSheet.getRows());
				
				system = new Label(1, reportSheet.getRows()-1,systemName.toUpperCase() ,
						createFormattedCell(10,  WritableFont.createFont("calibri"), true, false, null,  Border.ALL,
								null, null, Colour.WHITE, Colour.GRAY_50));
				reportSheet.mergeCells(3,reportSheet.getRows()-1, 5,reportSheet.getRows()-1);
				browserversion = new Label(3, reportSheet.getRows()-1,browserVersion.toUpperCase() ,
						createFormattedCell(10,  WritableFont.createFont("calibri"), true, false, null,  Border.ALL,
								null, null, Colour.WHITE, Colour.GRAY_50));
				reportSheet.addCell(Scenario);
				reportSheet.addCell(system);
			}else{
				if(browsercount==2){
					j=j+2;
				}
				Label result = new Label(j+5, 0,"RESULT" ,
						createFormattedCell(10,  WritableFont.createFont("calibri"), true, false, null,  Border.ALL,
								null, null, Colour.GRAY_80, Colour.LIGHT_TURQUOISE2));
				Label errlog= new Label(j+6, 0,"ERROR LOG" ,
						createFormattedCell(10,  WritableFont.createFont("calibri"), true, false, null,  Border.ALL,
								null, null, Colour.GRAY_80, Colour.LIGHT_TURQUOISE2));	
				Label Screenshots= new Label(j+7, 0,"SCREENSHOTS" ,
						createFormattedCell(10,  WritableFont.createFont("calibri"), true, false, null,  Border.ALL,
								null, null, Colour.GRAY_80, Colour.LIGHT_TURQUOISE2));
				Cell scenariocell = reportSheet.findCell(scenarioName);
				reportSheet.mergeCells(j+5,scenariocell.getRow(), j+7,scenariocell.getRow());
				browserversion = new Label(j+5,scenariocell.getRow(),browserVersion.toUpperCase() ,
						createFormattedCell(10,  WritableFont.createFont("calibri"), true, false, null,  Border.ALL,
								null, null, Colour.WHITE, Colour.GRAY_50));
				reportSheet.addCell(result);
				reportSheet.addCell(errlog);
				reportSheet.addCell(Screenshots);
				reportSheet.setColumnView(j+4,20 );
				reportSheet.setColumnView(j+5, 40);
				reportSheet.setColumnView(j+6, 45);
			}							
			reportSheet.addCell(browserversion);
			
			copy.write();
			copy.close();
			workbook.close();
		}catch(Exception e){
			e.printStackTrace();
		}
	}
	
	public boolean writeTestCaseWiseResults(int browserCount,int k,boolean fnconditionalflag,E2EDriver edata,long executionTime,String functionName,String systemName){
		Workbook workbook = null;
		WritableWorkbook copy = null;
		WorkbookSettings wbSettings = new WorkbookSettings();
		wbSettings.setLocale(new Locale("en", "ER"));
		boolean conditionflag=true;
		int column = 0,row = 0;
		Label status=null,ScenarioName,fnname;
		try{
			workbook = Workbook.getWorkbook(new File(MainController.outputFile), wbSettings);
			copy = Workbook.createWorkbook(new File(MainController.outputFile), workbook);			
			
			WritableSheet ScenarioSheet = copy.getSheet("TESTCASEWISE_REPORT");
			if(browserCount==0){
				ScenarioName= new Label(0, ScenarioSheet.getRows(), edata.getScenario() ,
						createFormattedCell(10,  WritableFont.createFont("calibri"), true, false, null,  Border.ALL,
								null, Alignment.LEFT, Colour.GRAY_80, Colour.LIGHT_TURQUOISE2));
				fnname= new Label(1, ScenarioSheet.getRows(), functionName ,
						createFormattedCell(10,  WritableFont.createFont("calibri"), true, false, null,  Border.ALL,
								null,  null, Colour.GRAY_80, Colour.LIGHT_TURQUOISE2));
				row=ScenarioSheet.getRows();
				column=2;
				ScenarioSheet.addCell(ScenarioName);
				ScenarioSheet.addCell(fnname);
				
				
			}else if(browserCount==1){
				Cell scenariocell=ScenarioSheet.findCell(edata.getScenario());
				row=scenariocell.getRow()+k;
				column=browserCount+3;
			}else if(browserCount==2){
				Cell scenariocell=ScenarioSheet.findCell(edata.getScenario());
				row=scenariocell.getRow()+k;
				column=browserCount+4;
			}
			if(fnconditionalflag){
				status= new Label(column, row,"PASS" ,
						createFormattedCell(10,  WritableFont.createFont("calibri"), true, false, null,  Border.ALL,
								null, null, Colour.GREEN, Colour.LIGHT_TURQUOISE2));
			}else{
				conditionflag=false;
				status= new Label(column, row,"FAIL" ,
						createFormattedCell(10,  WritableFont.createFont("calibri"), true, false, null,  Border.ALL,
								null, null, Colour.RED, Colour.LIGHT_TURQUOISE2));
			}
			
			Label time= new Label(column+1, row,Long.toString(executionTime) ,
					createFormattedCell(10,  WritableFont.createFont("calibri"), true, false, null,  Border.ALL,
							null, null, Colour.GRAY_80, Colour.LIGHT_TURQUOISE2));
			

			ScenarioSheet.addCell(status);
			ScenarioSheet.addCell(time);
			copy.write();
			copy.close();
			workbook.close();
		}catch(Exception e){
			e.printStackTrace();
		}
		return conditionflag;
	}
	
	
	public void writeDriverResult(boolean conditionflag,long executionTime,E2EDriver edata,String systemName){
		Workbook workbook = null;
		WritableWorkbook copy = null;
		WorkbookSettings wbSettings = new WorkbookSettings();
		wbSettings.setLocale(new Locale("en", "EN"));
		Label systemexetime = null,systemlabel=null;
		Cell systemLabelcell=null;
		int column=0;
		try{
			workbook = Workbook.getWorkbook(new File(MainController.outputFile), wbSettings);
			copy = Workbook.createWorkbook(new File(MainController.outputFile), workbook);
			WritableSheet driverSheet = copy.getSheet("DRIVER");
			Cell ScenarioCell=driverSheet.findCell(edata.getScenario());
			if(ScenarioCell==null){
				Label SNo = new Label(0, driverSheet.getRows(), String.valueOf(driverSheet.getRows()),
						createFormattedCell(11, WritableFont.createFont("calibri"), false, false, null, Border.ALL,
								null, null, Colour.GRAY_80, Colour.WHITE));
				Label Track = new Label(1, driverSheet.getRows(), edata.getTrack(),
						createFormattedCell(11, WritableFont.createFont("calibri"), false, false, null, Border.ALL,
								null, null, Colour.GRAY_80, Colour.WHITE));
				Label ComponentName = new Label(2, driverSheet.getRows(), edata.getScenario(),
						createFormattedCell(11, WritableFont.createFont("calibri"), false, false, null, Border.ALL,
								null, null, Colour.GRAY_80, Colour.WHITE));
				driverSheet.addCell(SNo);
				driverSheet.addCell(Track);
				driverSheet.addCell(ComponentName);
			}
			ScenarioCell=driverSheet.findCell(edata.getScenario());
			if(SystemConfig.systemFlag.equalsIgnoreCase("Specific Track")){
				column=3;
				systemName="";
			}else if(SystemConfig.systemFlag.equalsIgnoreCase("E2E Systems")){
				column=driverSheet.getColumns();
				systemLabelcell=driverSheet.findCell(systemName.toUpperCase()+" EXECUTION TIME(in Sec's)");
			}
			
			if(systemLabelcell==null){
				systemlabel = new Label(column, 0, systemName+" EXECUTION TIME(in Sec's)",
						createFormattedCell(11, WritableFont.createFont("calibri"), true, false, null, Border.ALL,
								null, null, Colour.WHITE, Colour.DARK_BLUE));
				driverSheet.addCell(systemlabel);
				driverSheet.setColumnView(column, 40);
				
			}
			systemLabelcell=driverSheet.findCell(systemName+" EXECUTION TIME(in Sec's)");
			if(!conditionflag){
				systemexetime= new Label(systemLabelcell.getColumn(), ScenarioCell.getRow(),Long.toString(executionTime) ,
						createFormattedCell(11, WritableFont.createFont("calibri"), false, false, null, Border.ALL,
								null, null, Colour.BRIGHT_GREEN, Colour.WHITE));
			}else{
				systemexetime= new Label(systemLabelcell.getColumn(), ScenarioCell.getRow(),Long.toString(executionTime),
						createFormattedCell(11, WritableFont.createFont("calibri"), false, false, null, Border.ALL,
								null, null, Colour.DARK_RED, Colour.WHITE));
			}
			driverSheet.addCell(systemexetime);			
			copy.write();
			copy.close();
			workbook.close();
		}catch(Exception e){
			e.printStackTrace();
		}
	}
	
	public boolean writeDetailedResults(int browserCount,int stepcount,ArrayList<SeleneseData> hd,E2EDriver edata,String functionName){
		Workbook workbook = null;
		WritableWorkbook copy = null;
		WorkbookSettings wbSettings = new WorkbookSettings();
		wbSettings.setLocale(new Locale("en", "ER"));
		Label Scenario,fnname,stepname,result = null,errorLog,screenshots;
		boolean flag = true;
		
		try{
			workbook = Workbook.getWorkbook(new File(MainController.outputFile), wbSettings);
			copy = Workbook.createWorkbook(new File(MainController.outputFile), workbook);
			
			/////////////////////// Detailed Report /////////////////////////
			WritableSheet reportSheet = copy.getSheet("DETAILED REPORT");
			Cell Scenariocell= reportSheet.findCell(edata.getScenario());
			int j=Scenariocell.getRow()+stepcount;
			for (int i = 0;hd !=null && i < hd.size(); i++) {
				
				if(browserCount==0){
					Scenario= new Label(0, reportSheet.getRows(),edata.getScenario() ,
							createFormattedCell(10,  WritableFont.createFont("calibri"), true, false, null,  Border.ALL,
									null,  Alignment.LEFT, Colour.GRAY_80, Colour.LIGHT_TURQUOISE2));
					fnname= new Label(1, reportSheet.getRows(),functionName ,
							createFormattedCell(10,  WritableFont.createFont("calibri"), true, false, null,  Border.ALL,
									null, null, Colour.GRAY_80, Colour.LIGHT_TURQUOISE2));
					stepname= new Label(2, reportSheet.getRows(),hd.get(i).getStep() ,
							createFormattedCell(10,  WritableFont.createFont("calibri"), true, false, null,  Border.ALL,
									null,  Alignment.LEFT, Colour.GRAY_80, Colour.LIGHT_TURQUOISE2));
					
					if(hd.get(i).getExstatus().equalsIgnoreCase("true")){					
						result= new Label(3, reportSheet.getRows(),"PASS",
								createFormattedCell(10,  WritableFont.createFont("calibri"), true, false, null,  Border.ALL,
										null, null, Colour.GREEN, Colour.LIGHT_TURQUOISE2));
					}else if(hd.get(i).getExstatus().equalsIgnoreCase("false")){
						flag = false;
						result= new Label(3, reportSheet.getRows(),"FAIL",
								createFormattedCell(10,  WritableFont.createFont("calibri"), true, false, null,  Border.ALL,
										null, null, Colour.RED, Colour.LIGHT_TURQUOISE2));
					}else {					
						result= new Label(3, reportSheet.getRows(),"SKIPPED",
								createFormattedCell(10,  WritableFont.createFont("calibri"), true, false, null,  Border.ALL,
										null, null, Colour.ORANGE, Colour.LIGHT_TURQUOISE2));
					}
						
					errorLog= new Label(4, reportSheet.getRows(),hd.get(i).getMessage(),
							createFormattedCell(10,  WritableFont.createFont("calibri"), true, false, null,  Border.ALL,
									null,  Alignment.LEFT, Colour.GRAY_80, Colour.LIGHT_TURQUOISE2));
					screenshots= new Label(5, reportSheet.getRows(),hd.get(i).getScreenshotPath(),
							createFormattedCell(10,  WritableFont.createFont("calibri"), true, false, null,  Border.ALL,
									null,  Alignment.LEFT, Colour.GRAY_80, Colour.LIGHT_TURQUOISE2));
					reportSheet.addCell(Scenario);
					reportSheet.addCell(fnname);
					reportSheet.addCell(stepname);				
					reportSheet.addCell(result);
					reportSheet.addCell(errorLog);
					reportSheet.addCell(screenshots);
				}else if(browserCount==1){
					
					if(hd.get(i).getExstatus().equalsIgnoreCase("true")){					
						result= new Label(browserCount+5, j,"PASS",
								createFormattedCell(10,  WritableFont.createFont("calibri"), true, false, null,  Border.ALL,
										null, null, Colour.GREEN, Colour.LIGHT_TURQUOISE2));
					}else if(hd.get(i).getExstatus().equalsIgnoreCase("false")){
						flag = false;
						result= new Label(browserCount+5, j,"FAIL",
								createFormattedCell(10,  WritableFont.createFont("calibri"), true, false, null,  Border.ALL,
										null, null, Colour.RED, Colour.LIGHT_TURQUOISE2));
					}else{					
						result= new Label(browserCount+5,  j,"SKIPPED",
								createFormattedCell(10,  WritableFont.createFont("calibri"), true, false, null,  Border.ALL,
										null, null, Colour.ORANGE, Colour.LIGHT_TURQUOISE2));
					}					
					errorLog= new Label(browserCount+6,j,hd.get(i).getMessage(),
							createFormattedCell(10,  WritableFont.createFont("calibri"), true, false, null,  Border.ALL,
									null,  Alignment.LEFT, Colour.GRAY_80, Colour.LIGHT_TURQUOISE2));
					screenshots= new Label(browserCount+7,  j,hd.get(i).getScreenshotPath(),
							createFormattedCell(10,  WritableFont.createFont("calibri"), true, false, null,  Border.ALL,
									null,  Alignment.LEFT, Colour.GRAY_80, Colour.LIGHT_TURQUOISE2));
					reportSheet.addCell(result);
					reportSheet.addCell(errorLog);
					reportSheet.addCell(screenshots);
					j++;
				}else{
					if(hd.get(i).getExstatus().equalsIgnoreCase("true")){					
						result= new Label(browserCount+7, j,"PASS",
								createFormattedCell(10,  WritableFont.createFont("calibri"), true, false, null,  Border.ALL,
										null, null, Colour.GREEN, Colour.LIGHT_TURQUOISE2));
					}else if(hd.get(i).getExstatus().equalsIgnoreCase("false")){
						flag = false;
						result= new Label(browserCount+7, j,"FAIL",
								createFormattedCell(10,  WritableFont.createFont("calibri"), true, false, null,  Border.ALL,
										null, null, Colour.RED, Colour.LIGHT_TURQUOISE2));
					}else{					
						result= new Label(browserCount+7,  j,"SKIPPED",
								createFormattedCell(10,  WritableFont.createFont("calibri"), true, false, null,  Border.ALL,
										null, null, Colour.ORANGE, Colour.LIGHT_TURQUOISE2));
					}					
					errorLog= new Label(browserCount+8,j,hd.get(i).getMessage(),
							createFormattedCell(10,  WritableFont.createFont("calibri"), true, false, null,  Border.ALL,
									null,  Alignment.LEFT, Colour.GRAY_80, Colour.LIGHT_TURQUOISE2));
					screenshots= new Label(browserCount+9,  j,hd.get(i).getScreenshotPath(),
							createFormattedCell(10,  WritableFont.createFont("calibri"), true, false, null,  Border.ALL,
									null,  Alignment.LEFT, Colour.GRAY_80, Colour.LIGHT_TURQUOISE2));
					reportSheet.addCell(result);
					reportSheet.addCell(errorLog);
					reportSheet.addCell(screenshots);
					j++;
				}
			}									
			copy.write();
			copy.close();
			workbook.close();
		}catch(Exception e){
			e.printStackTrace();
		}	
		return flag;
	}
		
}
