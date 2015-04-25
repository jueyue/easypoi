package org.jeecgframework.poi.test;

public class UtilTest {

    public static void main(String[] args) {
        String text = " {{   p    in    pList}}";
        text = text.replace("{{", "").replace("}}", "").replaceAll("\\s{1,}", " ").trim();
        System.out.println(text);
        System.out.println(text.length());
        
        
        
    }

}
