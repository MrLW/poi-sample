package com.lw.poi;

import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;

import javax.imageio.ImageIO;

import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
public class ImageTest {

	public static void main(String[] args) throws Exception {
		try {
			// 1、创建字节数组输出流
			ByteArrayOutputStream baos = new ByteArrayOutputStream();
			//2、读取一张图片放入ByteArrayOutStream
			BufferedImage bufferedImage = ImageIO.read(new File("D://timg.jpg"));
			// 3、将图缓冲写入字节数组中
			ImageIO.write(bufferedImage, "jpg", baos);
			/***************************************************/
			//4、创建工作铺
			HSSFWorkbook wb = new HSSFWorkbook();
			HSSFSheet imageSheet = wb.createSheet("imageSheet");
			// 5、获取画图的顶级管理器,一个sheet只能有一个
			HSSFPatriarch patriarch = imageSheet.createDrawingPatriarch();
			// 6、创建设置图片的属性
			HSSFClientAnchor anchor = new HSSFClientAnchor(0, 0, 255, 255, (short) 1, 1, (short) 5, 8);
			anchor.setAnchorType(3);
			// 7、插入图片
			patriarch.createPicture(anchor, wb.addPicture(baos.toByteArray(), HSSFWorkbook.PICTURE_TYPE_JPEG));
			// 8、创建Excel输出流
			FileOutputStream fos = new FileOutputStream("D:/测试.xls");
			// 9、写入Excel
			wb.write(fos);
			System.out.println("excel文件已经生成");
		} catch (Exception e) {
			// TODO: handle exception
			e.printStackTrace();
			System.out.println("Excel文件生产失败");
		}
	}
}
