package com.example.demo;

import java.io.InputStream;
import java.util.List;
import java.util.Map;

import org.json.JSONException;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import com.example.demo.util.ExcelUtil;
@RestController
@RequestMapping("/equiment")
public class text {

	@RequestMapping("/test1")
	public void test1() throws JSONException {

		System.out.println("123");
	}

	/**
	 * @return void
	 * @Description 产品导入
	 */

	@PostMapping("/uploadFile")
	public List<Map<String, String>> uploadMonitorItem(MultipartFile upfile, String providerId) throws Exception {

		InputStream in = null;
		List<Map<String, String>> listob = null;
		in = upfile.getInputStream();
		listob = ExcelUtil.readExcelByInputStream(in, providerId);

		return listob;
	}

}
