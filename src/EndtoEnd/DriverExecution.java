package EndtoEnd;


import java.awt.AWTException;
import java.awt.Robot;
import java.io.BufferedReader;
import java.io.File;
import java.io.IOException;
import java.io.InputStreamReader;
import java.net.MalformedURLException;
import java.net.URI;
import java.net.URISyntaxException;
import java.net.URL;
import java.nio.file.Paths;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import java.util.Locale;
import java.util.Set;
import java.util.TimeZone;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.format.Border;
import jxl.format.Colour;
import jxl.write.Label;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import mainController.MainController;

import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.ElementNotVisibleException;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebDriverException;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.interactions.HasInputDevices;
import org.openqa.selenium.interactions.Mouse;
import org.openqa.selenium.interactions.MoveTargetOutOfBoundsException;
import org.openqa.selenium.internal.Locatable;
import org.openqa.selenium.remote.CapabilityType;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.safari.SafariDriver;
import org.openqa.selenium.support.ui.ExpectedCondition;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.Wait;
import org.openqa.selenium.support.ui.WebDriverWait;

import com.gargoylesoftware.htmlunit.ElementNotFoundException;
import com.opera.core.systems.OperaDriver;
import com.thoughtworks.selenium.SeleniumException;

import configuration.SystemConfig;

import org.apache.commons.io.FileUtils;
import org.apache.commons.lang3.StringUtils;


public class DriverExecution {
	
	public String initDriver(String browserName, String iOSProxy) {
		String browserPath = new String("");
		int index = browserName.indexOf(" ");
		if (index != -1) {
			browserPath = (String) browserName.subSequence(index + 1,
					browserName.length());
		}
		if (browserName.toLowerCase().contains("firefox"))
			return initFirefoxDriver(browserPath);
		if (browserName.toLowerCase().contains("chrome"))
			return initChromeDriver(browserPath);
		if (browserName.toLowerCase().contains("iexplore"))
			return initIeDriver(browserPath);
		if (browserName.toLowerCase().contains("opera"))
			return initOperaDriver(browserPath);
		if (browserName.toLowerCase().contains("safari"))
			return initSafariDriver(browserPath);
		/*if (browserName.toLowerCase().contains("android"))
			return initAndroidDriver(browserPath);
		if (browserName.toLowerCase().contains("iphone")) 
			return initIPhoneDriver(browserPath,iOSProxy);*/
		return "UnRecognised Browser";
	}

	private WebDriver driver = null;

	public String initFirefoxDriver(String binaryPath) {
		if (!binaryPath.equals("")) {
			System.setProperty("webdriver.firefox.bin", binaryPath);
		}
		try {
			/*
			 * FirefoxProfile fp = new FirefoxProfile();
			 * fp.setEnableNativeEvents(true);
			 * fp.setPreference("network.proxy.type",
			 * "http://browserconfig.target.com/proxy-Global.pa");
			 */
			org.openqa.selenium.Proxy proxy = new org.openqa.selenium.Proxy();
			//proxy.setSslProxy("proxy-mdha.target.com:8080");
			proxy.setProxyAutoconfigUrl("http://browserconfig.target.com/proxy-Global.pac");
			DesiredCapabilities dc = DesiredCapabilities.firefox();
			dc.setCapability(CapabilityType.PROXY, proxy);
			this.driver = new FirefoxDriver(dc);
			this.driver.manage().deleteAllCookies();
			this.driver.manage().window().maximize();
		} catch (WebDriverException we) {
		} catch (Throwable th) {
		}
		String s = (String) ((JavascriptExecutor) this.driver).executeScript(
				"return navigator.userAgent;", new Object[0]);
		System.out.println(s);
		return getBrowserVersion("Firefox", s);
	}

	public String initIeDriver(String binaryPath) {
		System.setProperty("webdriver.ie.driver", "lib\\IEDriverServer.exe");
		String s;
		try {
			DesiredCapabilities iecapabilities = DesiredCapabilities.internetExplorer();
			iecapabilities.setCapability(InternetExplorerDriver.INTRODUCE_FLAKINESS_BY_IGNORING_SECURITY_DOMAINS,true);
			iecapabilities.setCapability("initialBrowserUrl", "https://tstgrscc.target.com");
			
			this.driver = new InternetExplorerDriver(iecapabilities);
			
			this.driver.manage().deleteAllCookies();
			this.driver.manage().window().maximize();
		} catch (WebDriverException we) {
			we.printStackTrace();
			System.out.println(we.getMessage());
			return "";
		} catch (Throwable th) {
			System.out.println(th.getMessage());
			return "";
		}
		s = (String) ((JavascriptExecutor) this.driver).executeScript(
				"return navigator.userAgent;", new Object[0]);
		return getBrowserVersion("InternetExplorer", s);

	}

	public String initChromeDriver(String binaryPath) {
		File dir = new File(".");
		String path = null;
		String OS = System.getProperty("os.name").toLowerCase();
		try {
			if (OS.indexOf("win") >= 0)
				path = dir.getCanonicalPath() + "\\lib\\chromedriver.exe";
			else if (OS.indexOf("mac") >= 0)
				path = "/usr/bin/chromedriver";
			if (!binaryPath.equals("")) {
				System.setProperty("chrome.binary", binaryPath);
			}
			System.setProperty("webdriver.chrome.driver", path);
		} catch (IOException e) {
			System.out.println(e.getMessage());
			return "";
		}
		try {
			//org.openqa.selenium.Proxy proxy = new org.openqa.selenium.Proxy();
			//proxy.setSslProxy("proxy-mdha.target.com:8080");
			
			//proxy.setProxyAutoconfigUrl("http://browserconfig.target.com/proxy-Global.pac");
			DesiredCapabilities dc = DesiredCapabilities.chrome();
			this.driver = new ChromeDriver(dc);
			this.driver.manage().deleteAllCookies();
			this.driver.manage().window().maximize();
		} catch (WebDriverException we) {
			we.printStackTrace();
			return "";
		} catch (Throwable th) {
			th.printStackTrace();
			return "";
		}
		String s = (String) ((JavascriptExecutor) this.driver).executeScript(
				"return navigator.userAgent;", new Object[0]);
		return getBrowserVersion("GoogleChrome", s);
	}

	public String initOperaDriver(String binaryPath) {
		if (!binaryPath.equals(""))
			System.setProperty("opera.binary", binaryPath);
		try {
			this.driver = new OperaDriver();
			this.driver.manage().window().maximize();
		} catch (WebDriverException we) {
			we.printStackTrace();
			return "";
		} catch (Throwable th) {
			th.printStackTrace();
			return "";
		}
		String s = (String) ((JavascriptExecutor) this.driver).executeScript(
				"return navigator.userAgent;", new Object[0]);
		return getBrowserVersion("Opera", s);
	}

	// Initiate Safari Driver...
	public String initSafariDriver(String binaryPath) {
		if (!binaryPath.equals("")) {
			System.setProperty("webdriver.safari.bin", binaryPath);
		}
		try {
			this.driver = new SafariDriver();
			this.driver.manage().deleteAllCookies();
			this.driver.manage().window().maximize();
		} catch (WebDriverException we) {
			we.printStackTrace();
		} catch (Throwable th) {
			th.printStackTrace();
		}
		String s = (String) ((JavascriptExecutor) this.driver).executeScript(
				"return navigator.userAgent;", new Object[0]);
		return getBrowserVersion("Safari", s);
	}

	


	// To Get Version of Browser we are executing...
	public String getBrowserVersion(String browserName,
			String JavaScriptUserAgent) {
		String browser_version = "";
		if (browserName.contains("InternetExplorer")) {
			browser_version = StringUtils.substringBetween(JavaScriptUserAgent,
					"MSIE", ";");
			browserName = "InternetExplorer " + browser_version;
		} else if (browserName.contains("Firefox")) {
			browser_version = StringUtils.substringAfterLast(
					JavaScriptUserAgent, "Firefox/");
			browserName = "Firefox " + browser_version;
		} else if (browserName.contains("GoogleChrome")) {
			browser_version = StringUtils.substringBetween(JavaScriptUserAgent,
					"Chrome/", " ");
			browserName = "GoogleChrome " + browser_version;
		} else if (browserName.contains("Opera")) {
			browser_version = StringUtils.substringBetween(JavaScriptUserAgent,
					"Opera/", " ");
			browserName = "Opera " + browser_version;
		} else if (browserName.contains("Safari")) {
			browser_version = StringUtils.substringBetween(JavaScriptUserAgent,
					"Version/", " ");
			browserName = "Safari " + browser_version;
		} else if (browserName.contains("Android")) {
			// browser_version = Build.VERSION.RELEASE;
			browserName = "Android " + browser_version;
		} else if (browserName.contains("IPhone")) {
			// browser_version = Build.VERSION.RELEASE;
			browserName = "IPhone " + browser_version;
		}
		return browserName;
	}

	// Capture Screenshot of Application alone
	public void captureScreen(String filename) {
		try {
			if (this.driver != null) {
				TakesScreenshot ts = (TakesScreenshot) this.driver;
				try {
					File scrFile = (File) ts.getScreenshotAs(OutputType.FILE);
					FileUtils.copyFile(scrFile, new File(filename));
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		}

		/*
		 * try { BufferedImage screencapture = new Robot()
		 * .createScreenCapture(new Rectangle(Toolkit
		 * .getDefaultToolkit().getScreenSize())); ImageIO.write(screencapture,
		 * "png", new File(filename)); } catch (Exception e) {
		 * e.printStackTrace(); }
		 */

	}
	
	public ArrayList<SeleneseData> triggerEXE(ArrayList<SeleneseData> hd,String filename){
		Process process = null;
		File dir = new File(".");
		BufferedReader input;
		String line;
		try{
			for (int i = 0; i < hd.size(); i++) {
				SeleneseData seleneseData = (SeleneseData) hd.get(i);
				try{
					process = new ProcessBuilder(dir.getCanonicalPath()+"\\exe\\"+filename,seleneseData.getCommand(),seleneseData.getValue()).start();
					process.waitFor();
					input = new BufferedReader(new InputStreamReader(process.getInputStream()));
					while ((line = input.readLine()) != null) {
						  System.out.println(line);
						  if(line.equalsIgnoreCase("true")){
							  seleneseData.setExstatus("true");
						  }else if(line.equalsIgnoreCase("false")){
							  seleneseData.setExstatus("false");
						  }else{
							  seleneseData.setMessage(line);
						  }
						}
				}catch(Exception e){
					e.printStackTrace();
				}								
				hd.remove(i);
				hd.add(i, seleneseData);
			}
		}catch(Exception e){
			e.printStackTrace();
		}	
		return hd;
	}
	
	public WebDriver loadApplication(String URL){
		URL url = null;
		// URL Parsing or URL Encoding				
				try {
					url = new URL(URL);
					URI uri = new URI(url.getProtocol(), url.getUserInfo(),
							url.getHost(), url.getPort(), url.getPath(),
							url.getQuery(), url.getRef());
					url = uri.toURL();
				} catch (MalformedURLException e2) {
					e2.printStackTrace();
				} catch (URISyntaxException e) {
					e.printStackTrace();
				}
				try {
					this.driver.get(url.toString());
					/*((JavascriptExecutor) driver).executeScript(
			                  "function pageloadingtime()"+
			                              "{"+
			                              "return 'Page has completely loaded'"+
			                              "}"+
			                  "return (window.onload=pageloadingtime());");*/
					ExpectedCondition<Boolean> expectation = new ExpectedCondition<Boolean>() {
						public Boolean apply(WebDriver driver) {
							return ((JavascriptExecutor) driver).executeScript(
									"return document.readyState").equals("complete");
						}
					};
					Wait<WebDriver> wait = new WebDriverWait(driver,30);				
				      try {
				              wait.until(expectation);
				      } catch(Throwable error) {
				             error.getStackTrace();
				      }		      
				} catch (Exception e) {
					System.out
							.println("Page could not be loaded within specified time: "
									+ e.getMessage());
				}
				return this.driver;
	}
	
	
	public ArrayList<SeleneseData> runserver(ArrayList<SeleneseData> hd, String outputPath,String browserVersion,String scenarioName) {
		
		String OS = System.getProperty("os.name").toLowerCase();
		File screenshotDirectory = null;
		if (OS.indexOf("win") >= 0)
			screenshotDirectory = new File(outputPath + "\\Screenshots\\" +  browserVersion);
		else if (OS.indexOf("mac") >= 0)
			screenshotDirectory = new File(outputPath + "/Screenshots/" + browserVersion);
		if (!screenshotDirectory.exists()) {
			screenshotDirectory.mkdirs();
		}
		for (int i = 0; hd!=null && i < hd.size() ; i++) {
			SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH-mm-ss.SSS");
			sdf.setTimeZone(TimeZone.getTimeZone("IST"));
			Date date = new Date();
			String screenShotPath = null;
			try {
				if (OS.indexOf("win") >= 0)
					screenShotPath = screenshotDirectory.getCanonicalPath()
							.toString() + "\\";
				else if (OS.indexOf("mac") >= 0)
					screenShotPath = screenshotDirectory.getCanonicalPath()
							.toString() + "/";
			} catch (IOException e1) {
				e1.printStackTrace();
			}
			screenShotPath = screenShotPath + sdf.format(date) + ".png";
			SeleneseData seleneseData = (SeleneseData) hd.get(i);			
			ArrayList<String> targetList = seleneseData.getTargetList();			
			WebElement element = null;
			WebElement element1 = null;
			int count = 0;							
				//while (count < targetList.size()) {
					seleneseData.setTarget((String) targetList.get(count));
					try {
						//element.add(count,  (WebElement) getElementFromLoc((String) targetList.get(count)));
						element  = getElementFromLoc((String) targetList.get(count));						
					} catch (Exception e) {
						System.out.println("Element: "
								+ (String) targetList.get(count) + "  --  "
								+ e.getMessage());
						seleneseData.setMessage(e.getMessage());
					}
					//count++;
				//}
				
				try {
					if(targetList.size() > 1){
						element1  = getElementFromLoc((String) targetList.get(1));
					}
				//	if (seleneseData.getExstatus().equalsIgnoreCase("false"))						
					seleneseData = executestep(element,element1,seleneseData, screenShotPath,browserVersion,scenarioName);
					if(seleneseData.getErrorFlag().equalsIgnoreCase("true") && seleneseData.getExstatus().equalsIgnoreCase("false")){
						seleneseData.setOverallStatus(false);								
						break;
					}else{
						seleneseData.setOverallStatus(true);
					}						
					System.out.println(seleneseData.getExstatus());
				} catch (NullPointerException e) {
					captureScreen(screenShotPath);
					seleneseData.setExstatus("false");
					seleneseData.setMessage("Element Not Found");
					e.printStackTrace();
				} catch (ElementNotVisibleException e) {
					captureScreen(screenShotPath);
					seleneseData.setExstatus("false");
					seleneseData.setMessage("Element not visible");
					e.printStackTrace();
				} catch (MoveTargetOutOfBoundsException e) {
					captureScreen(screenShotPath);
					seleneseData.setExstatus("false");
					seleneseData.setMessage("Element Cannot be scrolled into view");
					e.printStackTrace();
				} catch (ElementNotFoundException e) {
					captureScreen(screenShotPath);
					seleneseData.setExstatus("false");
					seleneseData.setMessage("Element Not Found");
					e.printStackTrace();
				} catch (SeleniumException e) {
					captureScreen(screenShotPath);
					seleneseData.setExstatus("false");
					seleneseData.setMessage(e.getMessage());
					e.printStackTrace();
				} catch (WebDriverException e) {
					captureScreen(screenShotPath);
					seleneseData.setExstatus("false");
					seleneseData.setMessage("Element Not Found");
					e.printStackTrace();
				} catch (AWTException e) {
					captureScreen(screenShotPath);
					seleneseData.setExstatus("false");
					seleneseData.setMessage("Key could not be pressed");
					e.printStackTrace();
				} catch (Exception e) {
					captureScreen(screenShotPath);
					seleneseData.setExstatus("false");
					seleneseData.setMessage(e.getMessage());
					e.printStackTrace();
				}
				hd.remove(i);
				hd.add(i, seleneseData);
				if(seleneseData.getCommand().equalsIgnoreCase("ifElementNotPresent") && seleneseData.getExstatus().equalsIgnoreCase("Skipped")){
					int j=i+1;										
					i+=((Integer.parseInt(seleneseData.getValue()))-1);
					for(;j<=i;j++){
						seleneseData = (SeleneseData) hd.get(j);
						System.out.println(seleneseData.getTarget());
						System.out.println(seleneseData.getCommand());
						seleneseData.setExstatus("Skipped");
						hd.remove(j);
						hd.add(j, seleneseData);
					}				
				}																	
		}
		try {
			
		} catch (Exception e) {
			System.out.println(e.getMessage());
			e.printStackTrace();
		}
		return hd;
	}		
		
		
	
	public WebElement getElementFromLoc(String target) throws Exception {
		WebElement webelement = null;
		StringBuffer element = new StringBuffer(target);
		try {
			if(target.contains("^")){
				String[] object =  target.split("^");				
				List<WebElement> listofwebelements = this.driver.findElements(By.xpath(object[0]));
				webelement =  listofwebelements.get(Integer.parseInt(object[1])-1);
			}else{
				if (target.startsWith("//")) {
					try {
						webelement = this.driver.findElement(By.xpath(element.toString()));
					} catch (ElementNotFoundException e) {
						e.printStackTrace();
						return null;
					}
				} else if (target.contains("xpath=")) {
					try {
						element.delete(0, 6);
						webelement =  this.driver.findElement(By.xpath(element
								.toString()));
					} catch (ElementNotFoundException e) {
						e.printStackTrace();
						return null;
					}
				} else if (target.contains("id=")) {
					try {
						element.delete(0, 3);
						webelement =  this.driver.findElement(By.id(element
								.toString()));
					} catch (ElementNotFoundException e) {
						e.printStackTrace();
						return null;
					}
				} else if (target.contains("name=")) {
					try {
						element.delete(0, 5);
						webelement = this.driver.findElement(By.name(element
								.toString()));
					} catch (ElementNotFoundException e) {
						e.printStackTrace();
						return null;
					}
				} else if (target.contains("css=")) {
					try {
						element.delete(0, 4);
						webelement =  this.driver.findElement(By.cssSelector(element
								.toString()));
					} catch (NoSuchElementException e) {
						e.printStackTrace();
						return null;
					}
				} else if (target.contains("link=")) {
					try {
						element.delete(0, 5);
						webelement =  this.driver.findElement(By.linkText(element
								.toString()));
					} catch (NoSuchElementException e) {
						e.printStackTrace();
						return null;
					}
				} else {
					webelement = null;
				}
			}						
			return  webelement;
		} catch (WebDriverException e) {
			System.out.println(e.getMessage());
			return null;
		}
	}		

	public SeleneseData executestep(WebElement element,WebElement element1,
			 SeleneseData hd, String outputPath,
			String browserVersion,String scenarioName) throws NullPointerException,
			AWTException, MoveTargetOutOfBoundsException,
			ElementNotVisibleException, ElementNotFoundException,
			WebDriverException, SeleniumException, Exception {
		SeleneseData hh = hd;
		System.out.println(hh.getTarget());
		String value, labelValue = null, text,line;
		String labelcell = null;
		StringBuffer label;
		Robot robot;
		Alert alert;
		File dir = new File(".");
		Workbook workbook = null;
		WritableWorkbook copy = null;
		WorkbookSettings wbSettings = new WorkbookSettings();
		wbSettings.setLocale(new Locale("en", "ER"));
		WritableSheet TestDataSheet = null;
		Sheet TestData=null;
		Cell scenarioCell,storelabelCell = null;
		BufferedReader input;
		Process process;
		String alertText = null;
		String mouseOnclickScript;
		String filepathGen=null;
		Actions action = new Actions(this.driver);
		JavascriptExecutor js = (JavascriptExecutor) driver;
		
		switch (hd.getCommand()) {
		case "webAuthPopup":
			String[] credentials=hd.getValue().split(",");
			process = new ProcessBuilder(dir.getCanonicalPath() + "\\lib\\WebAuthPopup.exe",credentials[0],credentials[1],credentials[2]).start();
			process.waitFor();
			break;
		case "triggerExe":
			process = new ProcessBuilder(dir.getCanonicalPath() + "\\exe\\"+hd.getValue()+".exe").start();
			process.waitFor();
			input = new BufferedReader(new InputStreamReader(process.getInputStream()));
			while ((line = input.readLine()) != null) {
				  System.out.println(line);
				}			
			hh.setExstatus("true");
			hh.setMessage(null);
			break;
		case "triggerVBS":
			String[] cred=hd.getValue().split(",");
			//process = new ProcessBuilder(dir.getCanonicalPath() + "\\exe\\"+cred[0]+".vbs",cred[1]).start();
			String path =dir.getCanonicalPath() + "\\lib\\"+cred[0]+".vbs";
			
			System.out.println(path);
			String[] command = {"cmd","/c", path,cred[1]};
			
			process = Runtime.getRuntime().exec(command);
			process.waitFor();
			input = new BufferedReader(new InputStreamReader(process.getInputStream()));
			while ((line = input.readLine()) != null) {
				  System.out.println(line);
				}			
			hh.setExstatus("true");
			hh.setMessage(null);
			break;
		case "ifFileDownload":
			/*String[] cred=hd.getValue().split(",");
			//process = new ProcessBuilder(dir.getCanonicalPath() + "\\exe\\"+cred[0]+".vbs",cred[1]).start();
			String path =dir.getCanonicalPath() + "\\lib\\"+cred[0]+".vbs";
			
			System.out.println(path);
			String[] command = {"cmd","/c", path,cred[1]};*/
			String exlName=hd.getValue();
			filepathGen=Paths.get(System.getProperty("user.home"), "Downloads").toString();
			File file1=new File(filepathGen +"\\"+exlName+".xlsx");
			boolean bool=true;
			int a=0;
			while(bool){
				if(file1.exists()){
					bool=false;
					break;
				}else{
					Thread.sleep(1000);
					a=a+1;
					if(a>=60){
						bool=false;
					}
				}
			}
			//input = new BufferedReader(new InputStreamReader(process.getInputStream()));
			//while ((line = input.readLine()) != null) {
			//	  System.out.println(line);
			//	}			
			hh.setExstatus("true");
			hh.setMessage(null);
			break;
		case "colorValidation":
			String color=element.getCssValue("color");
			if(color.regionMatches(true, 5, hd.getValue(), 5, color.length()-5)){
				if(SystemConfig.ScreenshotsFlag=="Both"){
					highlightElement(element);
					captureScreen(outputPath);
					unhighlightElement(element);
					hh.setScreenshotPath(outputPath);
				}
				if(SystemConfig.compatibilityFlag=="On")
					hh = setElementData(hh, element);
				hh.setExstatus("true");
				hh.setMessage(null);
			}else{
				if(SystemConfig.ScreenshotsFlag =="Fail" || SystemConfig.ScreenshotsFlag =="Both"){
					captureScreen(outputPath);
					hh.setScreenshotPath(outputPath);
				}						
				hh.setExstatus("false");	
			}
			break;
		case "dragAndDrop":			
			action.clickAndHold(element).release(element1).build().perform();
			Thread.sleep(5000L);
			if(SystemConfig.ScreenshotsFlag == "Fail" || SystemConfig.ScreenshotsFlag == "Both"){
				captureScreen(outputPath);
				hh.setScreenshotPath(outputPath);
			}
			hh.setExstatus("true");
			if(SystemConfig.compatibilityFlag=="On")
				hh = setElementData(hh, element1);
			hh.setMessage(null);
			break;
		case "httpsValidation":
			String URL=this.driver.getCurrentUrl();
			if(URL.startsWith("https"))
				hh.setExstatus("true");			
			else
				hh.setExstatus("false");
			break;
		case "switchTab":
			//new Actions(driver).sendKeys(driver.findElement(By.tagName("html")), Keys.CONTROL).sendKeys(driver.findElement(By.tagName("html")),Keys.NUMPAD2).build().perform();
			 ArrayList<String> tabs2 = new ArrayList<String> (driver.getWindowHandles());
			 this.driver.switchTo().window(tabs2.get(1));
			if(SystemConfig.ScreenshotsFlag == "Fail" || SystemConfig.ScreenshotsFlag == "Both"){
				captureScreen(outputPath);
				hh.setScreenshotPath(outputPath);
			}
			hh.setExstatus("true");
			
			break;
		case "switchParentTab":
			new Actions(driver).sendKeys(driver.findElement(By.tagName("html")), Keys.CONTROL).sendKeys(driver.findElement(By.tagName("html")),Keys.NUMPAD1).build().perform();
			if(SystemConfig.ScreenshotsFlag == "Fail" || SystemConfig.ScreenshotsFlag == "Both"){
				captureScreen(outputPath);
				hh.setScreenshotPath(outputPath);
			}
			hh.setExstatus("true");
			hh.setMessage(null);
			break;
		case "navigateBack":
			this.driver.navigate().back();
			if(SystemConfig.ScreenshotsFlag == "Fail" || SystemConfig.ScreenshotsFlag == "Both"){
				captureScreen(outputPath);
				hh.setScreenshotPath(outputPath);
			}
			hh.setExstatus("true");
			hh.setMessage(null);
			break;
		case "navigateForward":
			this.driver.navigate().forward();
			if(SystemConfig.ScreenshotsFlag == "Fail" || SystemConfig.ScreenshotsFlag == "Both"){
				captureScreen(outputPath);
				hh.setScreenshotPath(outputPath);
			}
			hh.setExstatus("true");
			hh.setMessage(null);
			break;
		case "refresh":
			this.driver.navigate().refresh();
			if(SystemConfig.ScreenshotsFlag == "Fail" || SystemConfig.ScreenshotsFlag == "Both"){
				captureScreen(outputPath);
				hh.setScreenshotPath(outputPath);
			}
			hh.setExstatus("true");
			hh.setMessage(null);
			break;
		case "sync":			
			try {
				/*((JavascriptExecutor) driver).executeScript(
		                  "function pageloadingtime()"+
		                              "{"+
		                              "return 'Page has completely loaded'"+
		                              "}"+
		                  "return (window.onload=pageloadingtime());");
				
				ExpectedConditions<Boolean> expectation = new ExpectedConditions<Boolean>() {
					public Boolean apply(WebDriver driver) {
						return ((JavascriptExecutor) driver).executeScript(
								"return document.readyState").equals("complete");
					}
				};*/
				//ExpectedConditions.presenceOfElementLocated(By.xpath(element.toString()));
				Wait<WebDriver> wait = new WebDriverWait(driver, Integer.parseInt(hd.getValue()));
			
				wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath(hd.getTarget())));
				hh.setExstatus("true");
				hh.setMessage(null);
				if(SystemConfig.ScreenshotsFlag=="Both"){
					captureScreen(outputPath);
					hh.setScreenshotPath(outputPath);
				}				
			} catch (Exception e) {
				e.printStackTrace()	;
				hh.setExstatus("false");
				if(SystemConfig.ScreenshotsFlag == "Fail" || SystemConfig.ScreenshotsFlag == "Both"){
					captureScreen(outputPath);
					hh.setScreenshotPath(outputPath);
				}
				hh.setMessage("Unable to load webpage within " + hd.getValue() + "Seconds");
			}
			break;
			
		case "ifFilePresentDelete":			
			try {
				String exlName1=hd.getValue();
				filepathGen=Paths.get(System.getProperty("user.home"), "Downloads").toString();
				File file=new File(filepathGen +"\\"+exlName1+".xlsx");
				
				if(file.exists()){
					
					file.delete();
				}
				hh.setExstatus("true");
				hh.setMessage(null);
				if(SystemConfig.ScreenshotsFlag=="Both"){
					captureScreen(outputPath);
					hh.setScreenshotPath(outputPath);
				}				
			} catch (Exception e) {
				e.printStackTrace()	;
				hh.setExstatus("false");
				if(SystemConfig.ScreenshotsFlag == "Fail" || SystemConfig.ScreenshotsFlag == "Both"){
					captureScreen(outputPath);
					hh.setScreenshotPath(outputPath);
				}
				hh.setMessage("Unable to load webpage within  " + hd.getValue() + "Seconds");
			}
			break;
		case "syncPageLoad":			
			try {				
				ExpectedCondition<Boolean> expectation = new ExpectedCondition<Boolean>() {
					public Boolean apply(WebDriver driver) {
						return ((JavascriptExecutor) driver).executeScript(
								"return document.readyState").equals("complete");
					}
				};
				Wait<WebDriver> wait = new WebDriverWait(driver,30);				
			      try {
			              wait.until(expectation);
			      } catch(Throwable error) {
			             error.getStackTrace();
			      }
				hh.setExstatus("true");
				hh.setMessage(null);
				if(SystemConfig.ScreenshotsFlag=="Both"){
					captureScreen(outputPath);
					hh.setScreenshotPath(outputPath);
				}				
			} catch (Exception e) {
				e.printStackTrace()	;
				hh.setExstatus("false");
				if(SystemConfig.ScreenshotsFlag == "Fail" || SystemConfig.ScreenshotsFlag == "Both"){
					captureScreen(outputPath);
					hh.setScreenshotPath(outputPath);
				}
				hh.setMessage("Unable to load webpage within 30 seconds");				
			}
			break;
		case "ifElementNotPresent":
			if (element != null  && element.isDisplayed()) {
			highlightElement(element);
				if(SystemConfig.ScreenshotsFlag == "Fail" || SystemConfig.ScreenshotsFlag == "Both")
					captureScreen(outputPath);
			unhighlightElement(element);
			if(SystemConfig.compatibilityFlag=="On")
				hh = setElementData(hh, element);			
			hh.setExstatus("true");
			hh.setMessage(null);
			hh.setScreenshotPath(outputPath);
			} else{
				if(SystemConfig.ScreenshotsFlag=="Both"){
					captureScreen(outputPath);
					hh.setScreenshotPath(outputPath);
				}									
			hh.setExstatus("Skipped");				
			hh.setMessage(null);
			}
			break;
		case "ifElementPresent":
			if (element != null && element.isDisplayed()) {
				if(SystemConfig.ScreenshotsFlag=="Both"){
					captureScreen(outputPath);
					hh.setScreenshotPath(outputPath);
				}									
			hh.setExstatus("Skipped");				
			hh.setMessage(null);
				
			
			} else{
				highlightElement(element);
				if(SystemConfig.ScreenshotsFlag == "Fail" || SystemConfig.ScreenshotsFlag == "Both")
					captureScreen(outputPath);
			unhighlightElement(element);
			if(SystemConfig.compatibilityFlag=="On")
				hh = setElementData(hh, element);			
			hh.setExstatus("true");
			hh.setMessage(null);
			hh.setScreenshotPath(outputPath);
			}
			break;
		case "verifyElementNotPresent":
			if (element != null) {
					if(SystemConfig.ScreenshotsFlag=="Both"){
						captureScreen(outputPath);
						hh.setScreenshotPath(outputPath);
					}		
					if(SystemConfig.compatibilityFlag=="On")
						hh = setElementData(hh, element);
				hh.setExstatus("true");
				hh.setMessage(null);
			} else {
				highlightElement(element);
					if(SystemConfig.ScreenshotsFlag == "Fail" || SystemConfig.ScreenshotsFlag == "Both")
						captureScreen(outputPath);
				unhighlightElement(element);				
				hh.setExstatus("false");
				hh.setMessage("Element is Present");
				hh.setScreenshotPath(outputPath);
			}
			break;
		case "open":
			loadApplication(hd.getValue());	
			if(SystemConfig.ScreenshotsFlag=="Fail" || SystemConfig.ScreenshotsFlag=="Both"){				
				captureScreen(outputPath);				
				hh.setScreenshotPath(outputPath);
			}
			hh.setExstatus("true");
			hh.setMessage(null);
			break;
		case "close":
			System.out.println("In Keyword Close");
			if(SystemConfig.ScreenshotsFlag=="Both"){
				captureScreen(outputPath);
				hh.setScreenshotPath(outputPath);
			}
			this.driver.close();
			hh.setExstatus("true");
			hh.setMessage(null);
			break;
		case "type":			
			if(element != null){				
				Thread.sleep(1000L);
				element.clear();
				element.sendKeys(new String[]{hd.getValue().trim()});
				if(SystemConfig.ScreenshotsFlag=="Both"){
					highlightElement(element);
					captureScreen(outputPath);
					unhighlightElement(element);
					hh.setScreenshotPath(outputPath);
				}
				if(SystemConfig.compatibilityFlag=="On")
					hh = setElementData(hh, element);				
				hh.setExstatus("true");
				hh.setMessage(null);
			}else{							
				if(SystemConfig.ScreenshotsFlag=="Fail" || SystemConfig.ScreenshotsFlag=="Both"){
					captureScreen(outputPath);
					hh.setScreenshotPath(outputPath);
				}						
				hh.setExstatus("false");			
			}
			break;
		case "selectWindow":
			Object[] list = this.driver.getWindowHandles().toArray();
			try {
				System.out.println(list.length);
				int index = Integer.parseInt(hd.getTarget());
				/*for (int i = 0; i < list.length; i++) {
					if (index == i) {
						hh.setExstatus("true");
						String windowHandle = (String) list[i];
						this.driver.switchTo().window(windowHandle);
						if(SystemConfig.ScreenshotsFlag=="Both"){
							captureScreen(outputPath);
							hh.setScreenshotPath(outputPath);
						}
					}
				}*/
				Thread.sleep(3000L);
				Set<String> winhandles = driver.getWindowHandles();
				Iterator<String> windowIterator =winhandles.iterator();

				System.out.println("window size = "+ winhandles.size());
				int i=0;
				while(windowIterator.hasNext()) {
					String windowHandle = windowIterator.next();					
					if(index == i){
						System.out.println("Value of i: "+ i +",index: "+ index);
						System.out.println(windowHandle);
						hh.setExstatus("true");
						hh.setMessage(null);
						driver.switchTo().window(windowHandle);
						break;
					}else{
						i++;
					}
					
				}
			} catch (NumberFormatException nfe) {
				nfe.printStackTrace();
				hh.setExstatus("false");
				if(SystemConfig.ScreenshotsFlag=="Fail" || SystemConfig.ScreenshotsFlag=="Both"){
					captureScreen(outputPath);
					hh.setScreenshotPath(outputPath);
				}
				hh.setMessage("Index not in numeric");				
			} catch (Exception e) {
				e.printStackTrace();
				hh.setExstatus("false");
				if(SystemConfig.ScreenshotsFlag=="Fail" || SystemConfig.ScreenshotsFlag=="Both"){
					captureScreen(outputPath);
					hh.setScreenshotPath(outputPath);
				}
				hh.setMessage("Could not switch to window");
			}
			break;
		case "storeAttributeValue":	
			if(element!=null){
				value = element.getAttribute("OldValue");
				if(SystemConfig.compatibilityFlag=="On")
					hh = setElementData(hh, element);
				try{
						/////////////Writing Stored Values to Label Provided///////////////////
						workbook =  Workbook.getWorkbook(new File(MainController.inputFile), wbSettings);
						copy = Workbook.createWorkbook(new File(MainController.inputFile), workbook);
						TestDataSheet = copy.getSheet("TEST DATA");
						scenarioCell = TestDataSheet.findCell(scenarioName);
						storelabelCell = TestDataSheet.findCell(hd.getValue());
						Label orderid = new Label(storelabelCell.getColumn(), scenarioCell.getRow(),value);			
						TestDataSheet.addCell(orderid);
						copy.write();
						copy.close();
						workbook.close();
						if(SystemConfig.ScreenshotsFlag=="Fail" || SystemConfig.ScreenshotsFlag=="Both"){
							captureScreen(outputPath);
							hh.setScreenshotPath(outputPath);
						}					
						hh.setExstatus("true");
						hh.setMessage(null);
						hh.setScreenshotPath(outputPath);	
				}catch(Exception e){
					e.printStackTrace();
				}
			}			
			break;			
		case "storeValue":
			try {
				labelcell=hd.getValue();				
				labelValue=null;
				captureScreen(outputPath);
				if(element!=null){
					if(SystemConfig.compatibilityFlag=="On")
						hh = setElementData(hh, element);
					labelValue=element.getText();
				}else{
					hh.setMessage("Element Not Found");
				}
				if(SystemConfig.ScreenshotsFlag=="Fail" || SystemConfig.ScreenshotsFlag=="Both"){
					captureScreen(outputPath);
					hh.setScreenshotPath(outputPath);
				}
			} catch (Exception e) {
				e.printStackTrace();
			}
			try{
				/////////////Writing Stored Values to Label Provided///////////////////
				workbook =  Workbook.getWorkbook(new File(MainController.inputFile), wbSettings);
				copy = Workbook.createWorkbook(new File(MainController.inputFile), workbook);
				TestDataSheet = copy.getSheet("TEST DATA");
				scenarioCell = TestDataSheet.findCell(scenarioName);
				storelabelCell = TestDataSheet.findCell(labelcell);				
				Label orderid = new Label(storelabelCell.getColumn(), scenarioCell.getRow(),labelValue);			
				TestDataSheet.addCell(orderid);
				copy.write();
				copy.close();
				workbook.close();
				
				workbook =  Workbook.getWorkbook(new File((dir.getCanonicalPath()+"\\Test_Report.xls")), wbSettings);
				copy = Workbook.createWorkbook(new File((dir.getCanonicalPath()+"\\Test_Report.xls")), workbook);
				TestDataSheet = copy.getSheet("OUTPUT_VALUES");
				scenarioCell = TestDataSheet.findCell(scenarioName);
				if(scenarioCell==null){
					Label SNO= new Label(0, TestDataSheet.getRows(),String.valueOf(TestDataSheet.getRows()),
							ExcelManipulation.createFormattedCell(10,  WritableFont.createFont("calibri"), true, false, null,  Border.ALL,
									null, null, Colour.GRAY_80, Colour.VERY_LIGHT_YELLOW));
					Label Scenario= new Label(1, TestDataSheet.getRows(),scenarioName ,
							ExcelManipulation.createFormattedCell(10,  WritableFont.createFont("calibri"), true, false, null,  Border.ALL,
									null, null, Colour.GRAY_80, Colour.VERY_LIGHT_YELLOW));			
					TestDataSheet.addCell(SNO);
					TestDataSheet.addCell(Scenario);			
					TestDataSheet.setColumnView(0, 15);
					TestDataSheet.setColumnView(1, 40);
				}
				scenarioCell = TestDataSheet.findCell(scenarioName);
				storelabelCell = TestDataSheet.findCell(labelcell);
				
				if(storelabelCell == null){
					Label TDlabel = new Label(TestDataSheet.getColumns(),0,labelcell,
							ExcelManipulation.createFormattedCell(10,  WritableFont.createFont("calibri"), true, false, null,  Border.ALL,
									null, null, Colour.GRAY_80, Colour.VERY_LIGHT_YELLOW));

					orderid = new Label(TestDataSheet.getColumns(), scenarioCell.getRow(),labelValue,
							ExcelManipulation.createFormattedCell(10,  WritableFont.createFont("calibri"), true, false, null,  Border.ALL,
									null, null, Colour.GRAY_80, Colour.VERY_LIGHT_YELLOW));
					TestDataSheet.addCell(TDlabel);
					TestDataSheet.addCell(orderid);
					TestDataSheet.setColumnView(TestDataSheet.getColumns(), 35);
					
				}else{
					orderid = new Label(storelabelCell.getColumn(), scenarioCell.getRow(),labelValue,
							ExcelManipulation.createFormattedCell(10,  WritableFont.createFont("calibri"), true, false, null,  Border.ALL,
									null, null, Colour.GRAY_80, Colour.VERY_LIGHT_YELLOW));
					TestDataSheet.addCell(orderid);
					TestDataSheet.setColumnView(storelabelCell.getColumn(), 35);
				}
				copy.write();
				copy.close();
				workbook.close();				
				hh.setExstatus("true");
				hh.setMessage(null);				
			}catch(Exception e){
				e.printStackTrace();
			}			
			break;
		
		case "retrieveValue":
			try{
				if(element != null){
				workbook =  Workbook.getWorkbook(new File(MainController.inputFile));				
				TestData = workbook.getSheet("TEST DATA");
				scenarioCell = TestData.findCell(scenarioName);
				if(hd.getValue().contains(",")){
					String[] valueArray=hd.getValue().split(",");
					storelabelCell = TestData.findCell(valueArray[0]);
					labelValue = TestData.getCell(storelabelCell.getColumn(), scenarioCell.getRow()).getContents().toString();
					labelValue = labelValue.substring(Integer.parseInt(valueArray[1]), Integer.parseInt(valueArray[2]));
					//String[] valueArray1 = labelValue.split("\\s");
					//labelValue = valueArray1[Integer.parseInt(valueArray[1])-1];
				}else{
					storelabelCell = TestData.findCell(hd.getValue());	
					labelValue = TestData.getCell(storelabelCell.getColumn(), scenarioCell.getRow()).getContents().toString();
				}																					
				Thread.sleep(1000L);
				element.clear();				
				element.sendKeys(new String[]{labelValue});
				workbook.close();
				}				
				hh.setExstatus("true");
				if(SystemConfig.compatibilityFlag=="On")
					hh = setElementData(hh, element);
				hh.setMessage(null);
				if(SystemConfig.ScreenshotsFlag=="Fail" || SystemConfig.ScreenshotsFlag=="Both"){
					captureScreen(outputPath);
					hh.setScreenshotPath(outputPath);
				}
			}catch(Exception e){
				e.printStackTrace();
			}
			break;
							
		case "verifyValue":
			try {
				if(element!=null){
					Thread.sleep(2000L);
					value = element.getAttribute("value");
					if (value.equals(hd.getValue())) {
						if(SystemConfig.compatibilityFlag=="On")
							hh = setElementData(hh, element);
						hh.setExstatus("true");
						hh.setMessage(null);
						if(SystemConfig.ScreenshotsFlag=="Both"){
							highlightElement(element);
							captureScreen(outputPath);
							unhighlightElement(element);
							hh.setScreenshotPath(outputPath);
						}
					} else {
						hh.setExstatus("false");
						hh.setMessage("Value is Different");
						if(SystemConfig.ScreenshotsFlag=="Fail" || SystemConfig.ScreenshotsFlag=="Both"){
							captureScreen(outputPath);
							hh.setScreenshotPath(outputPath);
						}
					}
				}				
			} catch (Exception localException2) {
				localException2.printStackTrace();
			}			
			break;
		
		case "isNotVisible":
			if (element != null && !element.isDisplayed()) {
				if(SystemConfig.ScreenshotsFlag=="Both"){
					highlightElement(element);
					captureScreen(outputPath);
					unhighlightElement(element);
					hh.setScreenshotPath(outputPath);
				}
				if(SystemConfig.compatibilityFlag=="On")
					hh = setElementData(hh, element);
				hh.setExstatus("true");
				hh.setMessage(null);
			} else {
				hh.setExstatus("false");
				hh.setMessage("Element is  Visible.");
				if(SystemConfig.ScreenshotsFlag=="Fail" || SystemConfig.ScreenshotsFlag=="Both"){
					highlightElement(element);
					captureScreen(outputPath);
					unhighlightElement(element);
					hh.setScreenshotPath(outputPath);
				}
			}
			break;
		case "isVisible":
			if (element != null && element.isDisplayed()) {
				if(SystemConfig.ScreenshotsFlag=="Both"){
					highlightElement(element);
					captureScreen(outputPath);
					unhighlightElement(element);
					hh.setScreenshotPath(outputPath);
				}
				if(SystemConfig.compatibilityFlag=="On")
					hh = setElementData(hh, element);
				hh.setExstatus("true");
				hh.setMessage(null);
			} else {
				hh.setExstatus("false");
				hh.setMessage("Element is not Visible.");
				if(SystemConfig.ScreenshotsFlag=="Fail" || SystemConfig.ScreenshotsFlag=="Both"){
					highlightElement(element);
					captureScreen(outputPath);
					unhighlightElement(element);
					hh.setScreenshotPath(outputPath);
				}
			}
			break;
			/////////////////////////////////////////////Done////////////////////////
		case "scrollUp":
			((JavascriptExecutor) driver).executeScript(
					"window.scrollTo(document.body.scrollHeight,0)", "");
			hh.setExstatus("true");
			hh.setMessage(null);
			if(SystemConfig.ScreenshotsFlag=="Both"){
				captureScreen(outputPath);
				hh.setScreenshotPath(outputPath);
			}
			break;
		case "scrollDown":
			((JavascriptExecutor) driver).executeScript(
					"window.scrollTo(0,document.body.scrollHeight)", "");
			hh.setExstatus("true");
			hh.setMessage(null);
			if(SystemConfig.ScreenshotsFlag=="Both"){
				captureScreen(outputPath);
				hh.setScreenshotPath(outputPath);
			}
			break;
		case "verifyText":
			if(element!=null){
				text = element.getText().trim() == null ? "" : element.getText()
						.trim();
				value = element.getAttribute("value") == null ? "" : element
						.getAttribute("value");
				if ((text.equals(hd.getValue().trim()))
						|| (value.equals(hd.getValue().trim()))) {
					if(SystemConfig.compatibilityFlag=="On")
						hh = setElementData(hh, element);
					hh.setExstatus("true");
					hh.setMessage(null);
					if(SystemConfig.ScreenshotsFlag=="Both"){
						highlightElement(element);
						captureScreen(outputPath);
						hh.setScreenshotPath(outputPath);
						unhighlightElement(element);
					}
				} else {
					hh.setExstatus("false");
					if (text.isEmpty())
						hh.setMessage("Text mismatch: " + value);
					else
						hh.setMessage("Text mismatch: " + text);
					if(SystemConfig.ScreenshotsFlag=="Fail" || SystemConfig.ScreenshotsFlag=="Both"){
						highlightElement(element);
						captureScreen(outputPath);
						hh.setScreenshotPath(outputPath);
						unhighlightElement(element);
					}
				}
			}			
			break;
		case "verifyNotText":
			if(element!=null){
				text = element.getText().trim() == null ? "" : element.getText()
						.trim();
				value = element.getAttribute("value") == null ? "" : element
						.getAttribute("value");
				if (!((text.equals(hd.getValue().trim())) || (value.equals(hd
						.getValue().trim())))) {
					if(SystemConfig.compatibilityFlag=="On")
						hh = setElementData(hh, element);
					hh.setExstatus("true");
					hh.setMessage(null);
					if(SystemConfig.ScreenshotsFlag=="Both"){
						highlightElement(element);
						captureScreen(outputPath);
						hh.setScreenshotPath(outputPath);
						unhighlightElement(element);
					}
				} else {
					hh.setExstatus("false");
					if (text.isEmpty())
						hh.setMessage("Text matches: " + value);
					else
						hh.setMessage("Text matches: " + text);
					if(SystemConfig.ScreenshotsFlag=="Fail" || SystemConfig.ScreenshotsFlag=="Both"){
						highlightElement(element);
						captureScreen(outputPath);
						hh.setScreenshotPath(outputPath);
						unhighlightElement(element);
					}
				}
			}			
			break;
		case "check":
			if(element != null){
				if (!element.isSelected())
					element.click();
				if(SystemConfig.ScreenshotsFlag=="Both"){
					highlightElement(element);
					captureScreen(outputPath);
					hh.setScreenshotPath(outputPath);
					unhighlightElement(element);
				}
				if(SystemConfig.compatibilityFlag=="On")
					hh = setElementData(hh, element);
				hh.setExstatus("true");
				hh.setMessage(null);
			}else{
				if(SystemConfig.ScreenshotsFlag=="Fail" || SystemConfig.ScreenshotsFlag=="Both"){
					captureScreen(outputPath);
					hh.setScreenshotPath(outputPath);
				}
				hh.setExstatus("false");
			}			
			break;
		case "unCheck":
			if(element != null){
				if (element.isSelected())
					element.click();
				if(SystemConfig.ScreenshotsFlag=="Both"){
					highlightElement(element);
					captureScreen(outputPath);
					hh.setScreenshotPath(outputPath);
					unhighlightElement(element);
				}
				if(SystemConfig.compatibilityFlag=="On")
					hh = setElementData(hh, element);
				hh.setExstatus("true");
				hh.setMessage(null);
			}else{
				if(SystemConfig.ScreenshotsFlag=="Fail" || SystemConfig.ScreenshotsFlag=="Both"){
					captureScreen(outputPath);
					hh.setScreenshotPath(outputPath);
				}
				hh.setExstatus("false");
			}
			break;
		case "acceptAlert":
			try{
				alert=this.driver.switchTo().alert();
				alert.accept();
				hh.setExstatus("true");
				hh.setMessage(null);
				if(SystemConfig.ScreenshotsFlag=="Both"){
					captureScreen(outputPath);
					hh.setScreenshotPath(outputPath);
				}				
			}catch(Exception e){
				hh.setExstatus("false");
				if(SystemConfig.ScreenshotsFlag=="Fail" || SystemConfig.ScreenshotsFlag=="Both"){
					captureScreen(outputPath);
					hh.setScreenshotPath(outputPath);
				}
			}
			break;
		case "getAlertText":
			try{
				alert=this.driver.switchTo().alert();
				alertText = alert.getText();
				hh.setExstatus("true");
				hh.setMessage(null);
				if(SystemConfig.ScreenshotsFlag=="Both"){
					captureScreen(outputPath);
					hh.setScreenshotPath(outputPath);
				}				
				/////////////Writing Stored Values to Label Provided///////////////////
				workbook =  Workbook.getWorkbook(new File(MainController.inputFile), wbSettings);
				copy = Workbook.createWorkbook(new File(MainController.inputFile), workbook);
				TestDataSheet = copy.getSheet("TEST DATA");
				scenarioCell = TestDataSheet.findCell(scenarioName);
				storelabelCell = TestDataSheet.findCell(hd.getValue());
				Label orderid = new Label(storelabelCell.getColumn(), scenarioCell.getRow(),alertText);			
				TestDataSheet.addCell(orderid);
				copy.write();
				copy.close();
				workbook.close();								
				hh.setExstatus("true");
				hh.setScreenshotPath(outputPath);
			}catch(Exception e){
				hh.setExstatus("false");
				if(SystemConfig.ScreenshotsFlag=="Fail" || SystemConfig.ScreenshotsFlag=="Both"){
					captureScreen(outputPath);
					hh.setScreenshotPath(outputPath);
				}
			}		
			break;
		case "highlight":
			highlightElement(element);
			try {
				Thread.sleep(2000L);
				captureScreen(outputPath);
				hh.setScreenshotPath(outputPath);
				unhighlightElement(element);
			} catch (Exception e) {
				e.printStackTrace();
			}
			if(SystemConfig.compatibilityFlag=="On")
				hh = setElementData(hh, element);
			break;
		case "pause":
			try {
				Thread.sleep(Long.parseLong(hd.getValue())*1000);
				hh.setExstatus("true");
				hh.setMessage(null);
				if(SystemConfig.ScreenshotsFlag=="Both"){
					captureScreen(outputPath);
					hh.setScreenshotPath(outputPath);
				}
			} catch (InterruptedException e) {
				hh.setExstatus("false");
				if(SystemConfig.ScreenshotsFlag=="Fail" || SystemConfig.ScreenshotsFlag=="Both"){
					captureScreen(outputPath);
					hh.setScreenshotPath(outputPath);
				}
			}
			break;
		case "verifyElementPresent":
			if(element != null){
				if(SystemConfig.ScreenshotsFlag=="Both"){
					highlightElement(element);
					captureScreen(outputPath);
					hh.setScreenshotPath(outputPath);
					unhighlightElement(element);					
				}
				if(SystemConfig.compatibilityFlag=="On")
					hh = setElementData(hh, element);
				hh.setExstatus("true");
				hh.setMessage(null);
			}else{
				if(SystemConfig.ScreenshotsFlag=="Fail" || SystemConfig.ScreenshotsFlag=="Both"){
					captureScreen(outputPath);
					hh.setScreenshotPath(outputPath);
				}
			}			
			break;
		case "select":
			if(element != null){
				if(SystemConfig.ScreenshotsFlag=="Both"){
					highlightElement(element);
					captureScreen(outputPath);
					hh.setScreenshotPath(outputPath);
					unhighlightElement(element);
				}
				if(SystemConfig.compatibilityFlag=="On")
					hh = setElementData(hh, element);
				hh.setExstatus("true");
				hh.setMessage(null);
				label = new StringBuffer(hd.getValue());
				new Select(element).selectByVisibleText(label.toString());						
			}else{
				if(SystemConfig.ScreenshotsFlag=="Fail" || SystemConfig.ScreenshotsFlag=="Both"){
					captureScreen(outputPath);
					hh.setScreenshotPath(outputPath);
				}
				hh.setMessage("Element Not Found");
				hh.setExstatus("false");
			}
			break;
		case "jsClick":
			try{								
				js.executeScript(hd.getValue());
				Thread.sleep(5000L);
				if(SystemConfig.ScreenshotsFlag=="Both"){					
					captureScreen(outputPath);
					hh.setScreenshotPath(outputPath);					
				}
				hh.setExstatus("true");
				hh.setMessage(null);
			}catch(Exception e){
				hh.setMessage("Element Not Found");
				hh.setExstatus("false");
				if(SystemConfig.ScreenshotsFlag=="Fail" || SystemConfig.ScreenshotsFlag=="Both"){
					captureScreen(outputPath);
					hh.setScreenshotPath(outputPath);
				}				
			}			
			break;
		case "executeJS":
			try{
				if(element!=null){
					System.out.println(hd.getValue());
					String date="arguments[0].setAttribute('value','"+hd.getValue()+"');";
					System.out.println(date);
					js.executeScript("arguments[0].setAttribute('value','');", element);
					js.executeScript(date, element);
					if(SystemConfig.compatibilityFlag=="On")
						hh = setElementData(hh, element);
					Thread.sleep(5000L);
					if(SystemConfig.ScreenshotsFlag=="Both"){					
						captureScreen(outputPath);
						hh.setScreenshotPath(outputPath);					
					}
					hh.setExstatus("true");
					hh.setMessage(null);
				}				
			}catch(Exception e){
				hh.setMessage("Element Not Found");
				hh.setExstatus("false");
				if(SystemConfig.ScreenshotsFlag=="Fail" || SystemConfig.ScreenshotsFlag=="Both"){
					captureScreen(outputPath);
					hh.setScreenshotPath(outputPath);
				}				
			}	
			break;
		case "mouseOnClickJS":
			try{
				mouseOnclickScript = "if(document.createEvent){var evObj = document.createEvent('MouseEvents');evObj.initEvent('onclick', true, false); arguments[0].dispatchEvent(evObj);} else if(document.createEventObject) { arguments[0].fireEvent('onclick');}";
				if(element != null){
					if(SystemConfig.ScreenshotsFlag=="Both"){
						highlightElement(element);
						captureScreen(outputPath);
						hh.setScreenshotPath(outputPath);
						unhighlightElement(element);
					}
					if(SystemConfig.compatibilityFlag=="On")
						hh = setElementData(hh, element);
					hh.setExstatus("true");
					hh.setMessage(null);
					js.executeScript(mouseOnclickScript, element);			
				}else{
					hh.setMessage("Element Not Found");
					hh.setExstatus("false");
					if(SystemConfig.ScreenshotsFlag=="Fail" || SystemConfig.ScreenshotsFlag=="Both"){
						captureScreen(outputPath);
						hh.setScreenshotPath(outputPath);
					}
				}							
			}catch(Exception e){
				hh.setMessage("Element Not Found");
				hh.setExstatus("false");
				if(SystemConfig.ScreenshotsFlag=="Fail" || SystemConfig.ScreenshotsFlag=="Both"){
					captureScreen(outputPath);
					hh.setScreenshotPath(outputPath);
				}
			}
			break;
		case "click":
			if(element != null){
				if(SystemConfig.ScreenshotsFlag=="Both"){
					highlightElement(element);
					captureScreen(outputPath);
					hh.setScreenshotPath(outputPath);
					unhighlightElement(element);
				}
				if(SystemConfig.compatibilityFlag=="On")
					hh = setElementData(hh, element);
				hh.setExstatus("true");
				hh.setMessage(null);
				element.click();				
			}else{
				hh.setMessage("Element Not Found");
				hh.setExstatus("false");
				if(SystemConfig.ScreenshotsFlag=="Fail" || SystemConfig.ScreenshotsFlag=="Both"){
					captureScreen(outputPath);
					hh.setScreenshotPath(outputPath);
				}
			}
			break;
		case "doubleClick":
			if(element != null){
				if(SystemConfig.ScreenshotsFlag=="Both"){
					highlightElement(element);
					captureScreen(outputPath);
					hh.setScreenshotPath(outputPath);
					unhighlightElement(element);
				}
				action.doubleClick(element).build().perform();
				if(SystemConfig.compatibilityFlag=="On")
					hh = setElementData(hh, element);
				hh.setExstatus("true");
				hh.setMessage(null);
			}else{
				hh.setMessage("Element Not Found");
				hh.setExstatus("false");
				if(SystemConfig.ScreenshotsFlag=="Fail" || SystemConfig.ScreenshotsFlag=="Both"){
					captureScreen(outputPath);
					hh.setScreenshotPath(outputPath);
				}
			}			
			break;
		case "clickAndWait":
			if(element != null){
				if(SystemConfig.ScreenshotsFlag=="Both"){
					highlightElement(element);
					captureScreen(outputPath);
					hh.setScreenshotPath(outputPath);
					unhighlightElement(element);
				}
				if(SystemConfig.compatibilityFlag=="On")
					hh = setElementData(hh, element);
				hh.setExstatus("true");
				hh.setMessage(null);
				element.click();				
				try{
					Thread.sleep(Long.parseLong(hd.getValue())*1000);
				}catch(Exception e){
					e.printStackTrace();
				}
			}else{
				hh.setMessage("Element Not Found");
				hh.setExstatus("false");
				if(SystemConfig.ScreenshotsFlag=="Fail" || SystemConfig.ScreenshotsFlag=="Both"){
					captureScreen(outputPath);
					hh.setScreenshotPath(outputPath);
				}
			}
			break;
		case "PressTab":
			robot = new Robot();
			robot.keyPress(9);
			robot.keyRelease(9);
			hh.setMessage(null);
			hh.setExstatus("true");
			hh.setScreenshotPath(outputPath);
			captureScreen(outputPath);
			break;
		case "Escape":
			robot = new Robot();
			robot.keyPress(27);
			robot.keyRelease(27);
			hh.setExstatus("true");
			hh.setMessage(null);
			hh.setScreenshotPath(outputPath);
			captureScreen(outputPath);
			break;
		case "PressEnter":
			if(element != null){
				if(SystemConfig.ScreenshotsFlag=="Both"){
					highlightElement(element);
					captureScreen(outputPath);
					hh.setScreenshotPath(outputPath);
					unhighlightElement(element);
				}
				if(SystemConfig.compatibilityFlag=="On")
					hh = setElementData(hh, element);
				hh.setExstatus("true");
				hh.setMessage(null);
				element.sendKeys(Keys.ENTER);							
			}else{
				hh.setMessage("Element Not Found");
				hh.setExstatus("false");
				if(SystemConfig.ScreenshotsFlag=="Fail" || SystemConfig.ScreenshotsFlag=="Both"){
					captureScreen(outputPath);
					hh.setScreenshotPath(outputPath);
				}
			}			
			break;
		case "selectFrame":
			if(hd.getValue().isEmpty()){
				hh.setMessage("FrameName have to be specifed in input sheet");
				hh.setExstatus("false");
				if(SystemConfig.ScreenshotsFlag=="Fail" || SystemConfig.ScreenshotsFlag=="Both"){
					captureScreen(outputPath);
					hh.setScreenshotPath(outputPath);
				}
			}else{
				if(hd.getValue().equalsIgnoreCase("defaultContent")){
					this.driver.switchTo().defaultContent();
				}else{
					this.driver.switchTo().frame(hd.getValue());
				}
				if(SystemConfig.ScreenshotsFlag=="Both"){					
					captureScreen(outputPath);
					hh.setScreenshotPath(outputPath);					
				}
				hh.setExstatus("true");
				hh.setMessage(null);
			}			
			break;
		case "mouseOver":
			if(element != null){
				hh.setExstatus("true");
				hh.setMessage(null);
				if(SystemConfig.compatibilityFlag=="On")
					hh = setElementData(hh, element);
				mouseOverOnElement(element);
				if(SystemConfig.ScreenshotsFlag=="Both"){
					highlightElement(element);
					captureScreen(outputPath);
					hh.setScreenshotPath(outputPath);
					unhighlightElement(element);
				}												
			}else{
				hh.setMessage("Element Not Found");
				hh.setExstatus("false");
				if(SystemConfig.ScreenshotsFlag=="Fail" || SystemConfig.ScreenshotsFlag=="Both"){
					captureScreen(outputPath);
					hh.setScreenshotPath(outputPath);
				}
			}
			break;
		case "mouseOverAndWait":
			if(element != null){
				hh.setExstatus("true");
				hh.setMessage(null);
				if(SystemConfig.compatibilityFlag=="On")
					hh = setElementData(hh, element);
				mouseOverOnElement(element);
				if(SystemConfig.ScreenshotsFlag=="Both"){
					highlightElement(element);
					captureScreen(outputPath);
					hh.setScreenshotPath(outputPath);
					unhighlightElement(element);
				}								
				try{
					Thread.sleep(Long.parseLong(hd.getValue()));
				}catch(Exception e){
					e.printStackTrace();
				}
			}else{
				hh.setMessage("Element Not Found");
				hh.setExstatus("false");
				if(SystemConfig.ScreenshotsFlag=="Fail" || SystemConfig.ScreenshotsFlag=="Both"){
					captureScreen(outputPath);
					hh.setScreenshotPath(outputPath);
				}
			}
			break;
		case "windowMaximize":
			hh.setExstatus("true");
			hh.setMessage(null);
			this.driver.manage().window().maximize();
			captureScreen(outputPath);
			hh.setScreenshotPath(outputPath);
			break;
		default:
			hh.setExstatus("false");
			hh.setScreenshotPath(outputPath);
			hh.setMessage("Command type not recognised");
			captureScreen(outputPath);
			break;
		}
		return hh;
	}

	

	public boolean highlightElement(WebElement elem) {
		try {
			JavascriptExecutor js = (JavascriptExecutor) this.driver;
			js.executeScript(
					"arguments[0].setAttribute('style', arguments[1]);",
					new Object[] { elem, "border: 1px solid green;" });
			return true;
		} catch (Exception e) {
		}
		return false;
	}

	public boolean unhighlightElement(WebElement elem) {
		try {
			JavascriptExecutor js = (JavascriptExecutor) this.driver;
			js.executeScript(
					"arguments[0].setAttribute('style', arguments[1]);",
					new Object[] { elem, "border: hidden;" });
			return true;
		} catch (Exception e) {
		}
		return false;
	}

	public void mouseOverOnElement(WebElement mouseOverElement) {
		try {
			if ((this.driver instanceof FirefoxDriver)) {
				Locatable button = (Locatable) mouseOverElement;
				Mouse mouse = ((HasInputDevices) this.driver).getMouse();
				mouse.mouseDown(button.getCoordinates());
			} else {
				Actions action = new Actions(this.driver);				
				action.moveToElement(mouseOverElement).build().perform();
				Thread.sleep(2000);
			}
		} catch (Exception e) {
			System.out.println(e.getMessage());
		}
	}

	public void focus(WebElement elem) {
		JavascriptExecutor js = (JavascriptExecutor) this.driver;
		try {
			js.executeScript("arguments[0].focus();", new Object[] { elem });
		} catch (Exception localException) {
		}
	}

	public SeleneseData setElementData(SeleneseData h, WebElement we) {
		h.setX(we.getLocation().getX());
		h.setY(we.getLocation().getY());		
		h.setHeight(we.getSize().getHeight());
		h.setWidth(we.getSize().getWidth());
		h.setColor(we.getCssValue("color"));
		h.setFontFamily(we.getCssValue("font-family"));
		h.setFontSize(we.getCssValue("font-size"));
		h.setFontStyle(we.getCssValue("font-style"));
		h.setFontWeight(we.getCssValue("font-weight"));
		return h;
	}

}
