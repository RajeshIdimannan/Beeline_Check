package EndtoEnd;

import java.util.ArrayList;
import java.util.List;

public class E2EDriver {
	private List<String> baseurl;
	private String Track;
	private String Scenario;
	private String Status;
	private List<String> browser;
	private int systemCount;	
	
	private List<String> systemName = new ArrayList<>();
	private ArrayList<eStoreData> eStoreData=null;

	
	
	
	public String getScenario() {
		return Scenario;
	}
	
	public void setScenario(String scenario) {
		Scenario = scenario;
	}
	public String getStatus() {
		return Status;
	}
	public void setStatus(String status) {
		Status = status;
	}
	
	
	public ArrayList<eStoreData> geteStoreData() {
		return eStoreData;
	}
	public void seteStoreData(ArrayList<eStoreData> eStoreData) {
		this.eStoreData = eStoreData;
	}
	
	public String getTrack() {
		return Track;
	}
	public void setTrack(String track) {
		Track = track;
	}
	
	
	
	public List<String> getBrowser() {
		return browser;
	}
	public void setBrowser(List<String> browser) {
		this.browser = browser;
	}
	public int getSystemCount() {
		return systemCount;
	}
	public void setSystemCount(int systemCount) {
		this.systemCount = systemCount;
	}
	public List<String> getBaseurl() {
		return baseurl;
	}
	public void setBaseurl(List<String> uRL) {
		this.baseurl = uRL;
	}

	public List<String> getSystemName() {
		return systemName;
	}

	public void setSystemName(List<String> systemName) {
		this.systemName = systemName;
	}

	
	

}
