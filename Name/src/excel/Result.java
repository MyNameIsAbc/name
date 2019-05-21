package excel;

public class Result {
	String name;
	String phone;
	String msg1;
	String msg2;
	public Result(String name, String phone, String msg1, String msg2) {
		super();
		this.name = name;
		this.phone = phone;
		this.msg1 = msg1;
		this.msg2 = msg2;
	}
	@Override
	public String toString() {
		return " [姓名 " + name + ", 手机号码  " + phone + ", " + msg1 + "," + msg2 + "]";
	}
	
}
