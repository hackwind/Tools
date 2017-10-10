package com.hjs.tools.ppt;

import java.awt.Color;
import java.awt.Dimension;
import java.awt.Graphics2D;
import java.awt.Image;
import java.awt.RenderingHints;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.PushbackInputStream;
import java.util.List;

import org.apache.poi.POIXMLDocument;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.sl.usermodel.Slide;
import org.apache.poi.sl.usermodel.SlideShow;
import org.apache.poi.sl.usermodel.TextRun;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.openxmlformats.schemas.drawingml.x2006.main.CTRegularTextRun;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextBody;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextCharacterProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextFont;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextParagraph;
import org.openxmlformats.schemas.presentationml.x2006.main.CTGroupShape;
import org.openxmlformats.schemas.presentationml.x2006.main.CTShape;
import org.openxmlformats.schemas.presentationml.x2006.main.CTSlide;
/**
 * 根据PPT生成一张预览图
 * @author Administrator
 *
 */
public class PPT2Image {
	/**
	 * 以下类方法中实现了将 2007和 2003的字体都转为宋体，防止中文出现乱码问题，源代码如下：
		类中的方法未判断 ppt的内容为空的时候，这个需要添加判断，只截取了第一个slide的图片
	 */
    private static Integer imgWidth=728;//默认宽度
    private static Integer imgHeight=409;//默认高度
    private static Integer padding = 20;
    public static void create2007PPTImage(InputStream in){
        try {
            XMLSlideShow xmlSlideShow=new XMLSlideShow(in);
            
            List<XSLFSlide> slides=xmlSlideShow.getSlides();
            Dimension dim = xmlSlideShow.getPageSize();
            imgWidth = dim.width;
            imgHeight = dim.height;
            BufferedImage img=new BufferedImage(imgWidth + padding * 3,
            		(int)(Math.ceil((slides.size() - 1) / 2.0f)) * (imgHeight / 2 + padding) //非第一张图片都是半尺寸
            		+ imgHeight + padding, //第一张图片是全尺寸
            		BufferedImage.TYPE_INT_RGB);
            Graphics2D graphics=img.createGraphics();
            graphics.setRenderingHint(RenderingHints.KEY_ANTIALIASING,  
                    RenderingHints.VALUE_ANTIALIAS_ON);  
            graphics.setRenderingHint(RenderingHints.KEY_RENDERING,  
                    RenderingHints.VALUE_RENDER_QUALITY);  
            graphics.setRenderingHint(RenderingHints.KEY_INTERPOLATION,  
                    RenderingHints.VALUE_INTERPOLATION_BICUBIC);  
            graphics.setRenderingHint(RenderingHints.KEY_FRACTIONALMETRICS,  
                    RenderingHints.VALUE_FRACTIONALMETRICS_ON);  
            graphics.setPaint(Color.WHITE);
            graphics.fillRect(0, 0, img.getWidth(), img.getHeight());
            int i = 0;
            int height = 0;
            for(XSLFSlide slide : slides) {
	            //设置字体为宋体，解决中文乱码问题
	            CTSlide rawSlide=slide.getXmlObject();
	            CTGroupShape gs = rawSlide.getCSld().getSpTree();
	            CTShape[] shapes = gs.getSpArray();
	            for (CTShape shape : shapes) {
	                CTTextBody tb = shape.getTxBody();
	                if (null == tb)
	                    continue;
	                CTTextParagraph[] paras = tb.getPArray();
	                CTTextFont font=CTTextFont.Factory.parse(
	                        "<xml-fragment xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">"+
	                        "<a:rPr lang=\"zh-CN\" altLang=\"en-US\" dirty=\"0\" smtClean=\"0\"> "+
	                            "<a:latin typeface=\"+mj-ea\"/> "+
	                          "</a:rPr>"+
	                        "</xml-fragment>");
	                for (CTTextParagraph textParagraph : paras) {
	                    CTRegularTextRun[] textRuns = textParagraph.getRArray();
	                    for (CTRegularTextRun textRun : textRuns) {
	                        CTTextCharacterProperties properties=textRun.getRPr();
	                        properties.setLatin(font);
	                    }
	                }
	            }
	            
	            System.out.println(i);
	            if(i == 0) {
	            	Image image = createSubImage(imgWidth,imgHeight,slide);
//	            	graphics.fill(new Rectangle2D.Float(padding, padding, imgWidth, imgHeight));
	            	graphics.drawImage(image,padding, padding, imgWidth, imgHeight, null);
	            	System.out.println("left:" + padding + ",top:" + padding + ",right:" + imgWidth + ",bottom:" + imgHeight);
	            	height += imgHeight + padding;
	            } else if(i % 2 == 1) {//左边
	            	Image image = createSubImage(imgWidth,imgHeight,slide);
//	            	graphics.fill(new Rectangle2D.Float(padding, height + padding , imgWidth / 2, imgHeight / 2));
	            	graphics.drawImage(image,padding, height + padding , imgWidth / 2, imgHeight / 2, null);
	            	System.out.println("left:" + padding + ",top:" + (height + padding) + ",right:" + imgWidth / 2  + ",bottom:" + imgHeight / 2);
	            } else {//右边
	            	Image image = createSubImage(imgWidth,imgHeight,slide);
//	            	graphics.fill(new Rectangle2D.Float(imgWidth / 2 + padding * 2, height + padding, imgWidth / 2, imgHeight / 2));
	            	graphics.drawImage(image,imgWidth / 2 + padding * 2, height + padding, imgWidth / 2, imgHeight / 2, null);
	            	System.out.println("left:" + (imgWidth / 2 + padding * 2) + ",top:" + (height + padding) + ",right:" + imgWidth / 2  + ",bottom:" + imgHeight / 2);
	            	height += imgHeight / 2 + padding;
	            }
//	            slide.draw(graphics);
	            
	            i++;
	            
            }
            FileOutputStream out = new FileOutputStream("E:/1.jpeg");   
            javax.imageio.ImageIO.write(img, "jpeg", out);   
            out.close();   
            
            System.out.println("生成缩略图成功!");
        } catch (Exception e) {
            e.printStackTrace();
        } 
    }
    
    private static Image createSubImage(int width,int height,XSLFSlide slide) {
    	BufferedImage subImg=new BufferedImage(width,height, BufferedImage.TYPE_INT_RGB);
        Graphics2D graphics=subImg.createGraphics();
        graphics.setRenderingHint(RenderingHints.KEY_ANTIALIASING,  
                RenderingHints.VALUE_ANTIALIAS_ON);  
        graphics.setRenderingHint(RenderingHints.KEY_RENDERING,  
                RenderingHints.VALUE_RENDER_QUALITY);  
        graphics.setRenderingHint(RenderingHints.KEY_INTERPOLATION,  
                RenderingHints.VALUE_INTERPOLATION_BICUBIC);  
        graphics.setRenderingHint(RenderingHints.KEY_FRACTIONALMETRICS,  
                RenderingHints.VALUE_FRACTIONALMETRICS_ON);  
        graphics.setPaint(Color.WHITE);
        slide.draw(graphics);
        return subImg;
    }
    
    
    public static void create2003PPTImage(InputStream in){
//        try {
//            SlideShow slideShow=new SlideShow(in);
//            
//            List<Slide> slides=slideShow.getSlides();
//            Slide slide=slides.get(0);
//            
//            TextRun[] textRuns=slide.getTextRuns();
//            for(TextRun tr:textRuns){
//               RichTextRun rt=tr.getRichTextRuns()[0];
//               rt.setFontName("宋体");
//            }
//           
//            BufferedImage img=new BufferedImage(imgWidth,imgHeight, BufferedImage.TYPE_INT_RGB);
//            Graphics2D graphics=img.createGraphics();
//            graphics.setPaint(Color.WHITE);
//            graphics.fill(new Rectangle2D.Float(0, 0, imgWidth, imgHeight));
//            slide.draw(graphics);
//           
//            FileOutputStream out = new FileOutputStream("E:/1.jpeg");   
//            javax.imageio.ImageIO.write(img, "jpeg", out);   
//            out.close();   
//            
//            System.out.println("缩略图成功!");
//        } catch (Exception e) {
//            e.printStackTrace();
//        } 
    }
    
    
    public static void createPPTImage(InputStream in){
//        try {
            if(!in.markSupported()){
                in=new BufferedInputStream(in);
            }
            if(in.markSupported()){
                in=new PushbackInputStream(in,8);
            }
//            if(POIFSFileSystem.hasPOIFSHeader(in)){//2003
//                create2003PPTImage(in);
//            }else { //if(POIXMLDocument.hasOOXMLHeader(in)){//2007
                create2007PPTImage(in);
//            }
//        } catch (IOException e) {
//            e.printStackTrace();
//        }
    }
    public static void main(String[] args) {
        try {
            createPPTImage(new  FileInputStream(new File("E:\\PPT模板.pptx")));
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
    }
}
