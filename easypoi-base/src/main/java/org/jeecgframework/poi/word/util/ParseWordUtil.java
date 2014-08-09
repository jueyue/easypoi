package org.jeecgframework.poi.word.util;

import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.lang.reflect.InvocationTargetException;
import java.util.List;
import java.util.Map;

import javax.imageio.ImageIO;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.jeecgframework.poi.util.POIPublicUtil;
import org.jeecgframework.poi.word.entity.WordImageEntity;
import org.jeecgframework.poi.word.entity.params.ExcelListEntity;


/**
 * 解析公告类
 * Excel 实体注解,暂时不支持图片
 * @author JueYue
 * @date 2014-7-24
 * @version 1.1
 */
public class ParseWordUtil {
	
	/**
	 * 解析数据
	 * 
	 * @Author JueYue
	 * @date 2013-11-16
	 * @return
	 */
	public static  Object getRealValue(String currentText,
			Map<String, Object> map) throws Exception {
		String params = "";
		while (currentText.indexOf("{{") != -1) {
			params = currentText.substring(currentText.indexOf("{{") + 2,
					currentText.indexOf("}}"));
			Object obj = getParamsValue(params.trim(),map);
			//判断图片或者是集合
			if(obj instanceof WordImageEntity||obj instanceof List ||obj instanceof ExcelListEntity){
				return obj;
			}else{
				currentText = currentText.replace("{{" + params + "}}",
						obj.toString());
			}
		}
		return currentText;
	}

	/**
	 * 获取参数值
	 * 
	 * @param params
	 * @param map
	 * @return
	 */
	private static  Object getParamsValue(String params, Map<String, Object> map) throws Exception {
		if (params.indexOf(".") != -1) {
			String[] paramsArr = params.split("\\.");
			return getValueDoWhile(map.get(paramsArr[0]), paramsArr, 1);
		}
		return map.containsKey(params) ? map.get(params) : "";
	}

	/**
	 * 通过遍历过去对象值
	 * 
	 * @param object
	 * @param paramsArr
	 * @param index
	 * @return
	 * @throws Exception
	 * @throws InvocationTargetException
	 * @throws IllegalAccessException
	 * @throws IllegalArgumentException
	 */
	@SuppressWarnings("rawtypes")
	public static  Object getValueDoWhile(Object object, String[] paramsArr,
			int index) throws Exception{
		if (object == null) {
			return "";
		}
		if(object instanceof WordImageEntity){
			return object;
		}
		if (object instanceof Map) {
			object = ((Map) object).get(paramsArr[index]);
		} else {
			object = POIPublicUtil.getMethod(paramsArr[index],
					object.getClass()).invoke(object, new Object[] {});
		}
		return (index == paramsArr.length - 1) ? (object == null ? "" : object) 
				: getValueDoWhile(object, paramsArr, ++index);
	}
	
	/**
	 * 返回流和图片类型
	 *@Author JueYue
	 *@date   2013-11-20
	 *@param obj
	 *@return  (byte[]) isAndType[0],(Integer)isAndType[1]
	 * @throws Exception 
	 */
	public static Object[] getIsAndType(WordImageEntity entity) throws Exception {
		Object[] result = new Object[2];
		String type;
		if(entity.getType().equals(WordImageEntity.URL)){
			ByteArrayOutputStream byteArrayOut = new ByteArrayOutputStream();
			BufferedImage bufferImg;
			String path = Thread.currentThread().getContextClassLoader().getResource("").toURI().getPath() +entity.getUrl();
			path = path.replace("WEB-INF/classes/","");
			path = path.replace("file:/","");
			bufferImg = ImageIO.read(
					new File(path));
			ImageIO.write(bufferImg,entity.getUrl().substring(entity.getUrl().indexOf(".")+1,entity.getUrl().length()),byteArrayOut);
			result[0] = byteArrayOut.toByteArray();
			type = entity.getUrl().split("/.")[ entity.getUrl().split("/.").length-1];
		}else{
			result[0] = entity.getData();
			type = POIPublicUtil.getFileExtendName(entity.getData());
		}
		result[1] = getImageType(type);
		return result;
	}
	
	private static Integer getImageType(String type) {
		if(type.equalsIgnoreCase("JPG")||type.equalsIgnoreCase("JPEG")){
			return XWPFDocument.PICTURE_TYPE_JPEG;
		}
		if(type.equalsIgnoreCase("GIF")){
			return XWPFDocument.PICTURE_TYPE_GIF;
		}
		if(type.equalsIgnoreCase("BMP")){
			return XWPFDocument.PICTURE_TYPE_GIF;
		}
		if(type.equalsIgnoreCase("PNG")){
			return XWPFDocument.PICTURE_TYPE_PNG;
		}
		return XWPFDocument.PICTURE_TYPE_JPEG;
	}


}
