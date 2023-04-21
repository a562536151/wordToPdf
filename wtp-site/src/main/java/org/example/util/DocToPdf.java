package org.example.util;

import com.aspose.words.Document;
import com.aspose.words.FontSettings;
import com.aspose.words.License;
import com.aspose.words.SaveFormat;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.Map;
import java.util.UUID;
/**
 * @author lucas
 * @program wtp-site
 * @description  word转pdf核心代码
 * @createDate 2020-03-04 23:05:31
 * 关注微信公众号"猿家"，600GJava架构师网盘资料等你下载
 * 这套代码我已经亲测可用了。若有特殊问题，请于公众号里随时与我联系。
 **/
public class DocToPdf {

    public static boolean getLicense() {
        boolean result = false;
        try {
            InputStream is = new FileInputStream(Constants.licensePath);
            License aposeLic = new License();
            aposeLic.setLicense(is);
            result = true;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return result;
    }

    public static  String DocToPdf(Map<String,Object> params,String documentXmlName,String documentDocName) throws Exception {
        try{
            if (!getLicense()) { // 验证License 若不验证则转化出的pdf文档会有水印产生
                return null;
            }

            String templateUUId = UUID.randomUUID().toString();
            String baseUrl=System.getProperty("java.io.tmpdir")+"/";
            String xml = baseUrl+templateUUId+".xml";
            XmlToDocx.toDocx(documentXmlName,Constants.docTemplatePath+documentDocName,xml,baseUrl+templateUUId+".docx",params);
            String docx = baseUrl+templateUUId+".docx";
            String pdf = baseUrl+templateUUId+".pdf";
            File file = new File(pdf); // 新建一个空白pdf文档
            FileOutputStream os = new FileOutputStream(file);
            Document doc = new Document(docx); // Address是将要被转化的word文档
            String osName = isWindows();
            String fonts [] = null;
            /**
             * 适配字体文件
             */
            /*if(osName.equals("linux")){
                fonts = new String[]{"/usr/share/fonts/chinese/","/usr/share/fonts/simsun/","/usr/share/fonts/black/","/usr/share/fonts/wingdings/"};
            }else if(osName.equals("mac")){
                fonts = new String[]{"C:\\Windows\\Fonts\\chinese","C:\\Windows\\Fonts\\simsun","C:\\Windows\\Fonts\\black","C:\\Windows\\Fontswingdings"};
            }*/

            fonts = new String[]{"C:\\Windows\\Fonts"};


            FontSettings.setFontsFolders(fonts, true);
            doc.save(os, SaveFormat.PDF);// 全面支持DOC, DOCX, OOXML, RTF HTML, OpenDocument, PDF,
            if(os!=null){
                os.close();
            }
            delete(xml,docx);//删除文件
            return pdf;
        }catch (Exception e){
            e.printStackTrace();
            return null;
        }

    }
    private static String isWindows() {
        String osName = System.getProperties().getProperty("os.name").toUpperCase();
        if(osName.indexOf("WINDOWS")!=-1){
            return "windows";
        }else if(osName.indexOf("MAC")!=-1){
            return "mac";
        }else{
            return "linux";
        }
    }
    public static void delete(String xmlPath,String wordPath){
        File file = new File(xmlPath);
        if(file.exists()){
            file.delete();
        }
        File wordFile = new File(wordPath);
        if(wordFile.exists()){
            wordFile.delete();
        }
    }

}
