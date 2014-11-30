package org.jeecgframework.poi.exception.word;

import org.jeecgframework.poi.exception.word.enmus.WordExportEnum;

/**
 * word导出异常
 * 
 * @author JueYue
 * @date 2014年8月9日 下午10:32:51
 */
public class WordExportException extends RuntimeException {

    private static final long serialVersionUID = 1L;

    public WordExportException() {
        super();
    }

    public WordExportException(String msg) {
        super(msg);
    }

    public WordExportException(WordExportEnum exception) {
        super(exception.getMsg());
    }

}
