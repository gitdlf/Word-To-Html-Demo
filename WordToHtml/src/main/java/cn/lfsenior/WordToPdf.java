package cn.lfsenior;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xwpf.converter.core.XWPFConverterException;
import org.apache.poi.xwpf.converter.pdf.PdfConverter;
import org.apache.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.junit.Test;

public class WordToPdf {
	/**
	 * word转换成pdf
	 */
	@Test
	public void wordToPdf(){
		try {
			String filePath = "E:/学习、练习数据文件夹/test/";
			String fileName = "SpringIOC解析.docx";
			String pdfName = "SpringIOC解析.pdf";
			//加载docx到XWPFDocument中
			FileInputStream in = new FileInputStream(new File(filePath+fileName));
			XWPFDocument document = new XWPFDocument(in);
			//创建pdfoptions
			PdfOptions options = PdfOptions.create();
			//将XWPFDocument转换成pdf
			FileOutputStream out = new FileOutputStream(new File(filePath+pdfName));
			PdfConverter.getInstance().convert(document, out, options);
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (XWPFConverterException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
}
