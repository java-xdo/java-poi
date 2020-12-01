package com.example.demo.util;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.PictureData;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtil {
	/**
	 * 读取 Excel文件内容
	 *
	 * @param inputstream 文件输入流
	 * @return
	 * @throws Exception
	 */
	public static List<Map<String, String>> readExcelByInputStream(InputStream inputstream, String providerId)
			throws Exception {
		String proFile = "D:/ruoyi/uploadPath";//文件存放的路径
		// 结果集
		List<Map<String, String>> list = new ArrayList<Map<String, String>>();

		XSSFWorkbook wb = new XSSFWorkbook(inputstream);
		//String filePath = ERPConfig.getProfile() + "/" + "pic/" + providerId + "/";//图片保存路径
		String filePath = proFile + "/" + "pic/" + providerId + "/";//图片保存路径
		final XSSFSheet sheet = wb.getSheetAt(0);// 得到Excel工作表对象

		Map<String, PictureData> map = ExcelImgUtil.getPictures(sheet);//获取图片和位置

		Map<String, String> pathMap = ExcelImgUtil.printImg(map, filePath);//写入图片，并返回图片路径，key：图片坐标，value：图片路径
		list = ExcelImgUtil.readData(sheet, pathMap,providerId);
		return list;
	}
}

