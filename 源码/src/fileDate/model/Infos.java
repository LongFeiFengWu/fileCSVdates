package fileDate.model;

public class Infos {

	private String mNum;
	private String tNum;
	private String pName;
	private double pMoney;
	private String payType;
	private String date;

	public String gettNum() {
		return tNum;
	}

	public void settNum(String tNum) {
		this.tNum = tNum;
	}
	public String getmNum() {
		return mNum;
	}

	public void setmNum(String mNum) {
		this.mNum = mNum;
	}

	public String getpName() {
		return pName;
	}

	public void setpName(String pName) {
		this.pName = pName;
	}

	public double getpMoney() {
		return pMoney;
	}

	public void setpMoney(double pMoney) {
		this.pMoney = pMoney;
	}

	public String getPayType() {
		return payType;
	}

	public void setPayType(String payType) {
		this.payType = payType;
	}

	public String getDate() {
		return date;
	}

	public void setDate(String date) {
		this.date = date;
	}

}
