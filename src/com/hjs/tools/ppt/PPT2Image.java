package com.hjs.tools.ppt;

import java.awt.Color;
import java.awt.Dimension;
import java.awt.Graphics2D;
import java.awt.Image;
import java.awt.RenderingHints;
import java.awt.image.BufferedImage;
import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.PushbackInputStream;
import java.util.List;

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
 * ����PPT����һ��Ԥ��ͼ
 * @author Administrator
 *
 */
public class PPT2Image {
	/**
	 * �����෽����ʵ���˽� 2007�� 2003�����嶼תΪ���壬��ֹ���ĳ����������⣬Դ�������£�
		���еķ���δ�ж� ppt������Ϊ�յ�ʱ�������Ҫ����жϣ�
	 */
    private static Integer imgWidth=728;//Ĭ�Ͽ��
    private static Integer imgHeight=409;//Ĭ�ϸ߶�
    private static Integer padding = 20;//�������߻��ж����ײ����
    private static Integer PIC_NUMBER = 2;//���˵�һ��ͼ��ȫͼ������ÿ�в���ͼƬ��
    private static Integer W_PADDING = 0;//ͼƬֱ�Ӽ��
    
    public static void main(String[] args) {
        try {
            createPPTImage(new  FileInputStream(new File("E:\\PPT\\PPTģ��.pptx")));
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
    }
    
    public static void createPPTImage(InputStream in){
//      try {
          if(!in.markSupported()){
              in=new BufferedInputStream(in);
          }
          if(in.markSupported()){
              in=new PushbackInputStream(in,8);
          }
//          if(POIFSFileSystem.hasPOIFSHeader(in)){//2003
//              create2003PPTImage(in);
//          }else { //if(POIXMLDocument.hasOOXMLHeader(in)){//2007
              create2007PPTImage(in);
//          }
//      } catch (IOException e) {
//          e.printStackTrace();
//      }
    }
    
    public static void create2007PPTImage(InputStream in){
        try {
            XMLSlideShow xmlSlideShow=new XMLSlideShow(in);
            
            List<XSLFSlide> slides=xmlSlideShow.getSlides();
            Dimension dim = xmlSlideShow.getPageSize();
            imgWidth = dim.width;
            imgHeight = dim.height;
            BufferedImage img=new BufferedImage(imgWidth + padding * 2,
            		(int)(Math.ceil((slides.size() - 1) / (float)PIC_NUMBER)) * (imgHeight / PIC_NUMBER) //�ǵ�һ��ͼƬ����1/4�ߴ�
            		+ imgHeight + padding * 2, //��һ��ͼƬ��ȫ�ߴ�
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
	            //��������Ϊ���壬���������������
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
	            int mod = i % PIC_NUMBER;
	            if(i == 0) {
	            	Image image = createSubImage(imgWidth,imgHeight,slide);
	            	graphics.drawImage(image,padding, padding, imgWidth, imgHeight, null);
	            	System.out.println("left:" + padding + ",top:" + padding + ",right:" + imgWidth + ",bottom:" + imgHeight);
	            	height += imgHeight + padding;
	            } else if(mod == 0){//����
	            	Image image = createSubImage(imgWidth,imgHeight,slide);
	            	graphics.drawImage(image,(PIC_NUMBER - 1) * imgWidth / PIC_NUMBER + W_PADDING * (PIC_NUMBER - 1) + padding, height + W_PADDING, imgWidth / PIC_NUMBER, imgHeight / PIC_NUMBER, null);
	            	System.out.println("left:" + ((PIC_NUMBER - 1) * imgWidth / PIC_NUMBER + W_PADDING * (PIC_NUMBER - 1) + padding) + ",top:" + (height + W_PADDING) + ",right:" + imgWidth / PIC_NUMBER  + ",bottom:" + imgHeight / PIC_NUMBER);
	            	height += imgHeight / PIC_NUMBER;
	            } else if(mod < PIC_NUMBER) {//��
	            	Image image = createSubImage(imgWidth,imgHeight,slide);
	            	graphics.drawImage(image,(mod - 1) * imgWidth / PIC_NUMBER + W_PADDING * (mod - 1) + padding, height + W_PADDING , imgWidth / PIC_NUMBER, imgHeight / PIC_NUMBER, null);
	            	System.out.println("left:" + ((mod - 1) * imgWidth / PIC_NUMBER + W_PADDING * (mod - 1) + padding) + ",top:" + (height + W_PADDING) + ",right:" + imgWidth / PIC_NUMBER  + ",bottom:" + imgHeight / PIC_NUMBER);
	            } 
	            
	            i++;
	            
            }
            FileOutputStream out = new FileOutputStream("E:/1.jpeg");   
            javax.imageio.ImageIO.write(img, "jpeg", out);   
            out.close();   
            
            System.out.println("��������ͼ�ɹ�!");
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
//               rt.setFontName("����");
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
//            System.out.println("����ͼ�ɹ�!");
//        } catch (Exception e) {
//            e.printStackTrace();
//        } 
    }
    
    
}
