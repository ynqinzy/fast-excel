package com.umon.fastexcel;


import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class TestExcel {

	public static void main(String[] args) {
		String filePath = "e://test_哈哈哈哈.xlsx";
		List<StudentVo> datas = new ArrayList<>();
		for (int i = 0; i < 100000; i++) {
			StudentVo vo = new StudentVo();
			vo.setStudentNo("00" + i);
			vo.setName("小明01");
			vo.setClassName("1班级");
			vo.setSex("男");
			vo.setTelephone("1565456564564");
			datas.add(vo);
		}
		long num = System.currentTimeMillis();
		// FastExcelUtils.createExcel(filePath, datas);
		try {
			List<StudentVo> list = FastExcelUtils.importExcel(StudentVo.class, filePath);
			System.out.println(System.currentTimeMillis() - num);
			for (StudentVo vo : list) {
				System.out.println(vo.getName());
			}
			System.out.println(list.size());
		} catch (IOException e) {
			e.printStackTrace();
		}

	}
}
