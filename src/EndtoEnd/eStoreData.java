package EndtoEnd;

import java.util.ArrayList;

public class eStoreData {
		
	private String fnname = null;
	
	private String parameter = null;
	private String systemName = null;
	private String message = null;
	private boolean exstatus = false;
	private ArrayList<SeleneseData> SeleneseData=null;
	private boolean estore;
	private boolean oms;
	
	public String getFnname() {
		return fnname;
	}
	public void setFnname(String fnname) {
		this.fnname = fnname;
	}
	public String getParameter() {
		return parameter;
	}
	public void setParameter(String parameter) {
		this.parameter = parameter;
	}
	public String getMessage() {
		return message;
	}
	public void setMessage(String message) {
		this.message = message;
	}
	public boolean isExstatus() {
		return exstatus;
	}
	public void setExstatus(boolean exstatus) {
		this.exstatus = exstatus;
	}
	public ArrayList<SeleneseData> getSeleneseData() {
		return SeleneseData;
	}
	public void setSeleneseData(ArrayList<SeleneseData> seleneseData) {
		SeleneseData = seleneseData;
	}
	public boolean isEstore() {
		return estore;
	}
	public void setEstore(boolean estore) {
		this.estore = estore;
	}
	public boolean isOms() {
		return oms;
	}
	public void setOms(boolean oms) {
		this.oms = oms;
	}
	public String getSystemName() {
		return systemName;
	}
	public void setSystemName(String systemName) {
		this.systemName = systemName;
	}
	
	
}
