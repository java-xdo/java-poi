package com.example.demo.util;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.POIXMLDocumentPart;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.PictureData;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTMarker;

public class ExcelImgUtil {
	private static int counter = 0;
	static String proFile = "D:/ruoyi/uploadPath";//文件存放的路径
	/**
	 * 获取图片和位置 (xlsx)
	 * 
	 * @param sheet
	 * @return
	 * @throws IOException
	 */
	public static Map<String, PictureData> getPictures(XSSFSheet sheet) throws IOException {
		Map<String, PictureData> map = new HashMap<String, PictureData>();
		List<POIXMLDocumentPart> list = sheet.getRelations();
		for (POIXMLDocumentPart part : list) {
			if (part instanceof XSSFDrawing) {
				XSSFDrawing drawing = (XSSFDrawing) part;
				List<XSSFShape> shapes = drawing.getShapes();
				for (XSSFShape shape : shapes) {
					XSSFPicture picture = (XSSFPicture) shape;
					XSSFClientAnchor anchor = picture.getPreferredSize();
					CTMarker marker = anchor.getFrom();
					String key = marker.getRow() + "-" + marker.getCol();
					byte[] data = picture.getPictureData().getData();
					map.put(key, picture.getPictureData());
				}
			}
		}
		return map;
	}

	public static Map<String, String> printImg(Map<String, PictureData> sheetList, String path) throws IOException {
		Map<String, String> pathMap = new HashMap<String, String>();
		Object[] key = sheetList.keySet().toArray();
		File f = new File(path);
		if (!f.exists()) {
			f.mkdirs(); // 创建目录
		}
		for (int i = 0; i < sheetList.size(); i++) {
			// 获取图片流
			PictureData pic = sheetList.get(key[i]);
			// 获取图片索引
			String picName = key[i].toString();
			// 获取图片格式
			String ext = pic.suggestFileExtension();
			String fileName = encodingFilename(picName);
			byte[] data = pic.getData();

			// 图片保存路径
			String imgPath = path + fileName + "." + ext;
			FileOutputStream out = new FileOutputStream(imgPath);

			imgPath = imgPath.substring(proFile.length(), imgPath.length());// 截取图片路径
			pathMap.put(picName, imgPath);
			out.write(data);
			out.close();
		}
		return pathMap;
	}

	private static final String encodingFilename(String fileName) {
		fileName = fileName.replace("_", " ");
		fileName = Md5Utils.hash(fileName + System.nanoTime() + counter++);
		return fileName;
	}

	/**
	 * 读取excel文字
	 * 
	 * Excel 07版本以上
	 * 
	 * @param sheet
	 */
	public static List<Map<String, String>> readData(XSSFSheet sheet, Map<String, String> map,String providerId) {
	
		List<Map<String, String>> newList = new ArrayList<Map<String, String>>();// 单行数据
		try {

			int rowNum = sheet.getLastRowNum() + 1;
			for (int i = 1; i < rowNum; i++) {// 从第三行开始读取数据,第一行是备注，第二行是标头

				Row row = sheet.getRow(i);// 得到Excel工作表的行
				if (row != null) {
					int col = row.getPhysicalNumberOfCells();
					// 单行数据
					
					Map<String, String> mapRes = new HashMap<String, String>();// 每格数据
					for (int j = 0; j < col; j++) {
						Cell cell = row.getCell(j);
						if (cell == null) {
							// arrayString.add("");
						} else if (cell.getCellType() == 0) {// 当时数字时的处理

							mapRes.put(getMapKey(j), new Double(cell.getNumericCellValue()).toString());
						} else {// 如果EXCEL表格中的数据类型为字符串型
							mapRes.put(getMapKey(j), cell.getStringCellValue().trim());

						}

					}

					if (i != 1) {// 不是标头列时，添加图片路径

						String path = map.get(i + "-9");
						mapRes.put(getMapKey(9), path);

					}
					mapRes.put("providerId", providerId);
					newList.add(mapRes);
				
				}

			}

		} catch (Exception e) {
		}
		return newList;
	}

	public static String getMapKey(int num) {
		String res = "";
		switch (num) {
		case 0:// 分类
			res = "thirdDictCode";
			break;
		case 1:// 产品名称
			res = "productName";
			break;
		case 2:// 规格型号
			res = "specification";
			break;
		case 3:// 计量单位
			res = "unit";
			break;
		case 4:// 风格
			res = "style";
			break;
		case 5:// 颜色
			res = "color";
			break;
		case 6:// 采购单价
			res = "purchasePrice";
			break;
		case 7:// 材质
			res = "material";
			break;
		case 8:// 备注
			res = "remark";
			break;

		case 9:// 产品图片
			res = "picture";
			break;

		default:
			break;
		}
		return res;
	}
}

