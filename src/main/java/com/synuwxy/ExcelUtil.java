package com.synuwxy;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class ExcelUtil {

		//工作薄
		private Workbook wb;
		//页面
		private Sheet sheet;
		//行
		private Row row;
		//单元格
		private Cell cell;
		
		public int getHSSFWorkbook(String path,List<String> fileList) {
			if(fileList.size() <= 0){
				return 0;
			}
			if(null == wb) {
				wb = new HSSFWorkbook();
			}
			sheet = wb.createSheet();
			int rowNum = 0;
			for (String string : fileList) {
				row = sheet.createRow(rowNum);
				cell = row.createCell(0);
				cell.setCellValue(rowNum+1);
				cell = row.createCell(1);
				cell.setCellValue(string);
				rowNum++;
			}
			File file;
			OutputStream os = null;
			try {
				file = new File(path+"/目录.xls");
				os = new FileOutputStream(file);
				wb.write(os);
			} catch (Exception e) {
				e.printStackTrace();
				return -1;
			}
			finally{
				try {
					if(null != os){
						os.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				}
		}
			return fileList.size();
		}
}
