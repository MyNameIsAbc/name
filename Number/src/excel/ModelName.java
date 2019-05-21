package excel;

public class ModelName {
	String name, mobile,examcode;

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public String getMobile() {
		return mobile;
	}

	public void setMobile(String mobile) {
		this.mobile = mobile;
	}

	public String getExamcode() {
		return examcode;
	}

	public void setExamcode(String examcode) {
		this.examcode = examcode;
	}

	@Override
	public String toString() {
		return "ModelName [name=" + name + ", mobile=" + mobile + ", examcode=" + examcode + "]";
	}
}
