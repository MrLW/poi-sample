package com.lw.poi;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.CellRangeAddress;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 *  单元格合并
 * @author lw
 */
public class MergedReginTest {

	public static void main(String[] args) throws Exception {
		FileOutputStream fos = new FileOutputStream("D:/测试.xls");
		
		Workbook wb = new HSSFWorkbook() ;
		
		Sheet mergedSheet = wb.createSheet("测试合并单元格");
		/**
		 *  设置单元格合并区域范围
		 */
		CellRangeAddress cr = new CellRangeAddress(10, 13, 13, 19);
		
		mergedSheet.addMergedRegion(cr) ;
		
		wb.write(fos);
		
		fos.close();
		
	}
}
