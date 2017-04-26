package com.lw.poi;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.CellRangeAddress;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 *  ��Ԫ��ϲ�
 * @author lw
 */
public class MergedReginTest {

	public static void main(String[] args) throws Exception {
		FileOutputStream fos = new FileOutputStream("D:/����.xls");
		
		Workbook wb = new HSSFWorkbook() ;
		
		Sheet mergedSheet = wb.createSheet("���Ժϲ���Ԫ��");
		/**
		 *  ���õ�Ԫ��ϲ�����Χ
		 */
		CellRangeAddress cr = new CellRangeAddress(10, 13, 13, 19);
		
		mergedSheet.addMergedRegion(cr) ;
		
		wb.write(fos);
		
		fos.close();
		
	}
}
