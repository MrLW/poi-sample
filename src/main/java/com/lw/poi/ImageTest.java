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
			// 1�������ֽ����������
			ByteArrayOutputStream baos = new ByteArrayOutputStream();
			//2����ȡһ��ͼƬ����ByteArrayOutStream
			BufferedImage bufferedImage = ImageIO.read(new File("D://timg.jpg"));
			// 3����ͼ����д���ֽ�������
			ImageIO.write(bufferedImage, "jpg", baos);
			/***************************************************/
			//4������������
			HSSFWorkbook wb = new HSSFWorkbook();
			HSSFSheet imageSheet = wb.createSheet("imageSheet");
			// 5����ȡ��ͼ�Ķ���������,һ��sheetֻ����һ��
			HSSFPatriarch patriarch = imageSheet.createDrawingPatriarch();
			// 6����������ͼƬ������
			HSSFClientAnchor anchor = new HSSFClientAnchor(0, 0, 255, 255, (short) 1, 1, (short) 5, 8);
			anchor.setAnchorType(3);
			// 7������ͼƬ
			patriarch.createPicture(anchor, wb.addPicture(baos.toByteArray(), HSSFWorkbook.PICTURE_TYPE_JPEG));
			// 8������Excel�����
			FileOutputStream fos = new FileOutputStream("D:/����.xls");
			// 9��д��Excel
			wb.write(fos);
			System.out.println("excel�ļ��Ѿ�����");
		} catch (Exception e) {
			// TODO: handle exception
			e.printStackTrace();
			System.out.println("Excel�ļ�����ʧ��");
		}
	}
}
