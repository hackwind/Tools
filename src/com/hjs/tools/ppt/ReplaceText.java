package com.hjs.tools.ppt;

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
 * @author Administrator
 *
 */

public class ReplaceText {
	//替换:   亮亮图文旗舰店  https://liangliangtuwen.tmall.com
	
	public static void main(String[] args) {
		String path = "E:\\PPT模板2.pptx";
		replace(path);
	}
	
	public static void replace(String path) {
		try { 
			XMLSlideShow ppt = new XMLSlideShow(new FileInputStream(path));

			Pattern pName = Pattern.compile("(亮亮图文旗舰店)");
			Matcher mName = null;
			
			Pattern pUrl = Pattern.compile("(https://liangliangtuwen.tmall.com)");
			Matcher mUrl = null;

			for (XSLFSlide slide : ppt.getSlides()) {
				for (XSLFShape shape : slide.getShapes()) {
					// System.out.println( shape.getShapeName() + " " +
					// shape.toString() ) ;

					if (shape instanceof XSLFTextShape) {
						XSLFTextShape txtshape = (XSLFTextShape) shape;
						System.out.println("XSLFTextShape" + ":" + shape.getShapeName() + ":" + txtshape.getText());
						//
						if (txtshape.getShapeName().equals("Text Box 14")) {
							txtshape.setText("");
						}

						mName = pName.matcher(txtshape.getText());
						txtshape.setText( mName.replaceAll("") ) ;
						
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
						XSLFConnectorShape connectorShape = (XSLFConnectorShape) shape;
						System.out.println("XSLFConnectorShape" + ":" + shape.getShapeName());
					} else if (shape instanceof XSLFPictureShape) {
						XSLFPictureShape picShape = (XSLFPictureShape) shape;
						System.out.println("XSLFPictureShape" + ":" + shape.getShapeName());
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
}
