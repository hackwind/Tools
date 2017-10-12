package com.hjs.tools.ppt;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFConnectorShape;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;

/**
 * 遍历PPT中每个页面文本，找到广告词语进行去除
 * 如替换:   亮亮图文旗舰店  https://liangliangtuwen.tmall.com
 * @author Administrator
 *
 */

public class ReplaceText {
	
	final static String DIRECTORY_PATH = "E://PPT/";
	final static String FILTER_FILE_NAME = "亮亮图文";
	
	public static void main(String[] args) {
		File file = new File(DIRECTORY_PATH);
		ergodic(file);
	}
	
	private static void ergodic(File parentFile) {
		if(parentFile.isDirectory()) {
			File[] files = parentFile.listFiles();
			for(File file :files) {
				ergodic(file);
			}
		} else {
			String newPath = renameName(parentFile.getAbsolutePath());
			replaceText(newPath);
			System.out.println("finish replace:" + parentFile + " to new :" + newPath);
		}
	}
	
	public static void replaceText(String path) {
		try { 
			XMLSlideShow ppt = new XMLSlideShow(new FileInputStream(path));

			Pattern pName = Pattern.compile("(亮亮图文旗舰店)");
			Matcher mName = null;
			
			Pattern pName2 = Pattern.compile("(亮亮图文)");
			Matcher mName2 = null;
			
			Pattern pUrl = Pattern.compile("(https://liangliangtuwen.tmall.com)");
			Matcher mUrl = null;

			for (XSLFSlide slide : ppt.getSlides()) {
				for (XSLFShape shape : slide.getShapes()) {
					// System.out.println( shape.getShapeName() + " " +
					// shape.toString() ) ;

					if (shape instanceof XSLFTextShape) {
						XSLFTextShape txtshape = (XSLFTextShape) shape;
//						System.out.println("XSLFTextShape" + ":" + shape.getShapeName() + ":" + txtshape.getText());
						//
						if (txtshape.getShapeName().equals("Text Box 14")) {
							txtshape.setText("");
						}

						mName = pName.matcher(txtshape.getText());
						txtshape.setText( mName.replaceAll("") ) ;
						
						mName2 = pName2.matcher(txtshape.getText());
						txtshape.setText( mName2.replaceAll("") ) ;
						
						mUrl = pUrl.matcher(txtshape.getText());
						txtshape.setText( mUrl.replaceAll("") ) ;
						// txtshape.setText( mmonth.replaceAll("09") ) ;
						// System.out.println(
						// txtshape.getTextParagraphs().get(0).getText() );
						// System.out.println(
						// txtshape.getTextParagraphs().get(0).getTextRuns().get(0).getText()
						// ) ;

						for (XSLFTextParagraph p : txtshape.getTextParagraphs()) {
							// System.out.println( p.getText() ) ;
							// System.out.println( p.getTextRuns().toString() )
							// ;
							for (XSLFTextRun textRun : p.getTextRuns()) {
								// System.out.println( textRun.getText() ) ;
							}
						}
					} else if (shape instanceof XSLFConnectorShape) {
//						XSLFConnectorShape connectorShape = (XSLFConnectorShape) shape;
//						System.out.println("XSLFConnectorShape" + ":" + shape.getShapeName());
					} else if (shape instanceof XSLFPictureShape) {
//						XSLFPictureShape picShape = (XSLFPictureShape) shape;
//						System.out.println("XSLFPictureShape" + ":" + shape.getShapeName());
					}
				}
			}

			FileOutputStream out = new FileOutputStream(path);
			ppt.write(out);
			out.close();
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	public static String renameName(String path) {
		File file = new File(path);
		if(!file.exists()) return "";
		String fileName = file.getName();
		if(fileName.contains(FILTER_FILE_NAME)) {
			System.out.println("rename " + path);
			String newPath = path.replace(FILTER_FILE_NAME, "").replaceAll("-", "");
			File newFile = new File(newPath);
			if(file.renameTo(newFile)){
				return newPath;
			} else {
				return "";
			}
		} else {
			return path;
		}
	}
}
