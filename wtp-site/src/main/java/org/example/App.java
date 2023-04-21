package org.example;

import org.example.util.*;

import javax.annotation.Resource;
import java.io.File;
import java.util.HashMap;
import java.util.Map;


public class App 
{
    /**
     * 注意事项
     * 1：如果转换之后有乱码现象，极大的概率是你的word当中用了系统不支持的字体。
     * @param args
     * @throws Exception
     */

    public static void main( String[] args ) throws Exception {
        Map<String,Object> params = new HashMap<String,Object>();
        params.put("number","20230101");
        params.put("company","北京国家电网");
        params.put("name","北京北京北京北京");
        params.put("money","2333333333");
        DocxToXmlConverter docxToXmlConverter = new DocxToXmlConverter();
        docxToXmlConverter.docxToXml();
        String pdfPath = DocToPdf.DocToPdf(params, "document.xml", "template.docx");

        //pdfPath便是最后在系统磁盘里生成的绝对路径地址了，后续业务如果需要，可以读取这个路径进行把文件上传到云空间里或者其他地方。
        //如果上传到云空间了，记得把这个文件在系统磁盘删掉，避免占用系统磁盘。
        System.out.println("pdf已生成，路径是："+pdfPath);
    }
}
