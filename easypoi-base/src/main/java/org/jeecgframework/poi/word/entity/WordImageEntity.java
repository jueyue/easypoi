package org.jeecgframework.poi.word.entity;
/**
 * word导出,图片设置和图片信息
 * @author JueYue
 * @date 2013-11-17
 * @version 1.0
 */
public class WordImageEntity {
	
	public static String URL = "url";
	public static String Data = "data";
	/**
	 * 图片输入方式
	 */
	private String type = URL;
	/**
	 * 图片宽度
	 */
	private int width;
	//图片高度
	private int height;
	//图片地址
	private String url;
	//图片信息
	private byte[] data;
	
	public String getType() {
		return type;
	}

	public void setType(String type) {
		this.type = type;
	}

	public int getWidth() {
		return width;
	}

	public void setWidth(int width) {
		this.width = width;
	}

	public int getHeight() {
		return height;
	}

	public void setHeight(int height) {
		this.height = height;
	}

	public String getUrl() {
		return url;
	}

	public void setUrl(String url) {
		this.url = url;
	}

	public byte[] getData() {
		return data;
	}

	public void setData(byte[] data) {
		this.data = data;
	}

}
