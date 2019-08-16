package com.synuwxy;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.List;


public class WordUtil {

    public int getXWPFDocument(String path,List<String> fileList){
		if(fileList.size() <= 0){
			return 0;
		}
        XWPFDocument doc = new XWPFDocument();
		XWPFTable table= doc.createTable(fileList.size(), 2);
		int rowNum = 0;
		for (String string : fileList) {
			table.getRow(rowNum).getCell(0).setText(rowNum+1+"");
			table.getRow(rowNum).getCell(1).setText(string);
			rowNum++;
		}
		File file;
		OutputStream os = null;
		try {
			file = new File(path + "/目录.doc");
			os = new FileOutputStream(file);
			doc.write(os);
		} catch (Exception e) {
			e.printStackTrace();
			return -1;
		}
		finally{
			try {
				if(null != os) {
				os.close();
				}
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
		return fileList.size();
	}

}
