package org.jeecgframework.poi.word.entity;

import java.io.InputStream;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlToken;
import org.openxmlformats.schemas.drawingml.x2006.main.CTNonVisualDrawingProps;
import org.openxmlformats.schemas.drawingml.x2006.main.CTPositiveSize2D;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.CTInline;

/**
 * 扩充document,修复图片插入失败问题问题
 * 
 * @author JueYue
 * @date 2013-11-20
 * @version 1.0
 */
public class JeecgXWPFDocument extends XWPFDocument {

	public JeecgXWPFDocument() {
		super();
	}

	public JeecgXWPFDocument(OPCPackage opcPackage) throws Exception {
		super(opcPackage);
	}

	public JeecgXWPFDocument(InputStream in) throws Exception {
		super(in);
	}
	
	 public void createPicture(String blipId,int id, int width, int height) {  
	        final int EMU = 9525;  
	        width *= EMU;  
	        height *= EMU;  
	        //String blipId = getAllPictures().get(id).getPackageRelationship().getId();  
	        CTInline inline = createParagraph().createRun().getCTR().addNewDrawing().addNewInline();  
	  
	        String picXml = "" +  
	                "<a:graphic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" +  
	                "   <a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" +  
	                "      <pic:pic xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" +  
	                "         <pic:nvPicPr>" +  
	                "            <pic:cNvPr id=\"" + id + "\" name=\"Generated\"/>" +  
	                "            <pic:cNvPicPr/>" +  
	                "         </pic:nvPicPr>" +  
	                "         <pic:blipFill>" +  
	                "            <a:blip r:embed=\"" + blipId + "\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"/>" +  
	                "            <a:stretch>" +  
	                "               <a:fillRect/>" +  
	                "            </a:stretch>" +  
	                "         </pic:blipFill>" +  
	                "         <pic:spPr>" +  
	                "            <a:xfrm>" +  
	                "               <a:off x=\"0\" y=\"0\"/>" +  
	                "               <a:ext cx=\"" + width + "\" cy=\"" + height + "\"/>" +  
	                "            </a:xfrm>" +  
	                "            <a:prstGeom prst=\"rect\">" +  
	                "               <a:avLst/>" +  
	                "            </a:prstGeom>" +  
	                "         </pic:spPr>" +  
	                "      </pic:pic>" +  
	                "   </a:graphicData>" +  
	                "</a:graphic>";  
	  
	        //CTGraphicalObjectData graphicData = inline.addNewGraphic().addNewGraphicData();  
	        XmlToken xmlToken = null;  
	        try {  
	            xmlToken = XmlToken.Factory.parse(picXml);  
	        } catch(XmlException xe) {  
	            xe.printStackTrace();  
	        }  
	        inline.set(xmlToken);  
	        //graphicData.set(xmlToken);  
	  
	        inline.setDistT(0);  
	        inline.setDistB(0);  
	        inline.setDistL(0);  
	        inline.setDistR(0);  
	  
	        CTPositiveSize2D extent = inline.addNewExtent();  
	        extent.setCx(width);  
	        extent.setCy(height);  
	  
	        CTNonVisualDrawingProps docPr = inline.addNewDocPr();  
	        docPr.setId(id);  
	        docPr.setName("Picture " + id);  
	        docPr.setDescr("Generated");  
	    }  

	public void createPicture(XWPFRun run,String blipId, int id, int width, int height) {
		final int EMU = 9525;
		width *= EMU;
		height *= EMU;
		CTInline inline = run.getCTR()
				.addNewDrawing().addNewInline();

		String picXml = ""
				+ "<a:graphic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">"
				+ "   <a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">"
				+ "      <pic:pic xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">"
				+ "         <pic:nvPicPr>" + "            <pic:cNvPr id=\""
				+ id
				+ "\" name=\"Generated\"/>"
				+ "            <pic:cNvPicPr/>"
				+ "         </pic:nvPicPr>"
				+ "         <pic:blipFill>"
				+ "            <a:blip r:embed=\""
				+ blipId
				+ "\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"/>"
				+ "            <a:stretch>"
				+ "               <a:fillRect/>"
				+ "            </a:stretch>"
				+ "         </pic:blipFill>"
				+ "         <pic:spPr>"
				+ "            <a:xfrm>"
				+ "               <a:off x=\"0\" y=\"0\"/>"
				+ "               <a:ext cx=\""
				+ width
				+ "\" cy=\""
				+ height
				+ "\"/>"
				+ "            </a:xfrm>"
				+ "            <a:prstGeom prst=\"rect\">"
				+ "               <a:avLst/>"
				+ "            </a:prstGeom>"
				+ "         </pic:spPr>"
				+ "      </pic:pic>"
				+ "   </a:graphicData>" + "</a:graphic>";

		XmlToken xmlToken = null;
		try {
			xmlToken = XmlToken.Factory.parse(picXml);
		} catch (XmlException xe) {
			xe.printStackTrace();
		}
		inline.set(xmlToken);

		inline.setDistT(0);
		inline.setDistB(0);
		inline.setDistL(0);
		inline.setDistR(0);

		CTPositiveSize2D extent = inline.addNewExtent();
		extent.setCx(width);
		extent.setCy(height);

		CTNonVisualDrawingProps docPr = inline.addNewDocPr();
		docPr.setId(id);
		docPr.setName("Picture " + id);
		docPr.setDescr("Generated");
	}

}
