package cn.lfsenior;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.converter.PicturesManager;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.apache.poi.hwpf.usermodel.PictureType;
import org.apache.poi.xwpf.converter.core.FileImageExtractor;
import org.apache.poi.xwpf.converter.core.FileURIResolver;
import org.apache.poi.xwpf.converter.xhtml.XHTMLConverter;
import org.apache.poi.xwpf.converter.xhtml.XHTMLOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.junit.Test;
import org.w3c.dom.Document;

/**
 * 将word以上版本转成Html
 * 
 * @author LFSenior
 *
 */
public class WordToHtml {
	/**
	 * 将Word2007+转成Html
	 * 
	 * @throws Exception
	 */
	@Test
	public void word2007ToHtml() throws Exception {
		String filePath = "E:/学习、练习数据文件夹/test/";
		String fileName = "SpringIOC解析.docx";
		String htmlName = "SpringIOC解析.html";
		final String file = filePath + fileName;
		File f = new File(file);
		if (!f.exists()) {
			System.out.println("Sorry File does not Exists!");
		} else {
			/* 判断是否为docx文件 */
			if (f.getName().endsWith(".docx") || f.getName().endsWith(".DOCX")) {
				// 1)加载word文档生成XWPFDocument对象
				FileInputStream in = new FileInputStream(f);
				XWPFDocument document = new XWPFDocument(in);
				// 2)解析XHTML配置（这里设置IURIResolver来设置图片存放的目录）
				File imageFolderFile = new File(filePath);
				XHTMLOptions options = XHTMLOptions.create().URIResolver(new FileURIResolver(imageFolderFile));
				options.setExtractor(new FileImageExtractor(imageFolderFile));
				options.setIgnoreStylesIfUnused(false);
				options.setFragment(true);
				// 3)将XWPFDocument转换成XHTML
				FileOutputStream out = new FileOutputStream(new File(filePath + htmlName));
				XHTMLConverter.getInstance().convert(document, out, options);
				//也可以使用字符数组流获取解析的内容
				//ByteArrayOutputStream baos = new ByteArrayOutputStream(); 
				//XHTMLConverter.getInstance().convert(document, baos, options);  
				//String content = baos.toString();
				//System.out.println(content);
				//baos.close();
			} else {
				System.out.println("Enter only as MS Office 2007+ files");
			}
		}
	}

	/**
	 * word2003-2007转换成html
	 * @throws Exception
	 */
	@Test
	public void wordToHtml() throws Exception {
		String filePath = "E:/学习、练习数据文件夹/test/";
		String fileName = "SpringIOC解析2003.doc";
		String htmlName = "SpringIOC解析2003.html";
		final String imagePath = filePath + "/image/";
		final String file = filePath + fileName;
		InputStream input = new FileInputStream(new File(file));
		HWPFDocument wordDocument = new HWPFDocument(input);
		WordToHtmlConverter wordToHtmlConverter = new WordToHtmlConverter(
				DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument());
		//设置图片存储位置
		wordToHtmlConverter.setPicturesManager(new PicturesManager() {
			
			public String savePicture(byte[] content, PictureType pictureType, String suggestedName, float widthInches,
					float heightInches) {
				File imgPath=new File(imagePath);
				if (!imgPath.exists()) {//目录不存在则创建目录
					imgPath.mkdirs();
				}
				File file = new File(imagePath+suggestedName);
				try {
					FileOutputStream os = new FileOutputStream(file);
					os.write(content);
					os.close();
				} catch (FileNotFoundException e) {
					e.printStackTrace();
				} catch (IOException e) {
					e.printStackTrace();
				}
				return imagePath+suggestedName;
			}
		});
		
		//解析word文档
		wordToHtmlConverter.processDocument(wordDocument);
		Document htmlDocument = wordToHtmlConverter.getDocument();
		File htmlFile = new File(filePath+htmlName);
		FileOutputStream outStream = new FileOutputStream(htmlFile);
	    //也可以使用字符数组流获取解析的内容
		//ByteArrayOutputStream baos = new ByteArrayOutputStream(); 
		//OutputStream outStream = new BufferedOutputStream(baos);
		DOMSource domSource = new DOMSource(htmlDocument);
		StreamResult streamResult = new StreamResult(outStream);
		TransformerFactory factory = TransformerFactory.newInstance();
		Transformer serializer = factory.newTransformer();
		serializer.setOutputProperty(OutputKeys.ENCODING,"utf-8");
		serializer.setOutputProperty(OutputKeys.INDENT, "yes");
		serializer.setOutputProperty(OutputKeys.METHOD, "html");
		serializer.transform(domSource, streamResult);
		//也可以使用字符数组流获取解析的内容
		//String content = baos.toString();
		//System.out.println(content);
		//baos.close();
		outStream.close();
	}
}
