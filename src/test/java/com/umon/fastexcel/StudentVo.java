package com.umon.fastexcel;

import com.umon.fastexcel.annotation.ExcelCell;
import com.umon.fastexcel.annotation.ExcelSheet;

@ExcelSheet
public class StudentVo {

	@ExcelCell(name = "班级名称", sn = 0)
	private String className;

	@ExcelCell(name = "学号", sn = 1)
	private String studentNo;

	@ExcelCell(name = "姓名", sn = 2)
	private String name;

	@ExcelCell(name = "性别", sn = 3)
	private String sex;

	@ExcelCell(name = "手机号", sn = 4)
	private String telephone;

	private String birthday;


	public String getStudentNo() {
		return studentNo;
	}

	public void setStudentNo(String studentNo) {
		this.studentNo = studentNo;
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public String getSex() {
		return sex;
	}

	public void setSex(String sex) {
		this.sex = sex;
	}

	public String getBirthday() {
		return birthday;
	}

	public void setBirthday(String birthday) {
		this.birthday = birthday;
	}

	public String getTelephone() {
		return telephone;
	}

	public void setTelephone(String telephone) {
		this.telephone = telephone;
	}

	public String getClassName() {
		return className;
	}

	public void setClassName(String className) {
		this.className = className;
	}


}
