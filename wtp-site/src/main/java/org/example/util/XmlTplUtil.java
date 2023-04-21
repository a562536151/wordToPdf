package org.example.util;
import freemarker.template.Configuration;
import freemarker.template.Template;
import freemarker.template.TemplateException;
import java.io.File;
import java.io.IOException;
import java.io.Writer;
import java.util.Map;

import static org.example.util.Constants.docXmlTemplatePath;

/**
 * @description  word动态填充工具类
 **/
public class XmlTplUtil {

    private static XmlTplUtil tplm = null;
    private Configuration cfg = null;

    private XmlTplUtil() {
        cfg = new Configuration();
        try {
            cfg.setDirectoryForTemplateLoading(new File(docXmlTemplatePath));
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static Template getTemplate(String name) throws IOException {
        if (tplm == null) {
            tplm = new XmlTplUtil();
        }
        Template template = tplm.cfg.getTemplate(name,"utf-8");
        return template;
    }

    /**
     * @param templatefile 模板文件
     * @param param        需要填充的内容
     * @param out          填充完成输出的文件
     * @throws IOException
     * @throws TemplateException
     */
    public static void process(String templatefile, Map param, Writer out) throws IOException, TemplateException {
        // 获取模板
        Template template = XmlTplUtil.getTemplate(templatefile);
        template.setOutputEncoding("utf-8");
        // 合并数据
        template.process(param, out);
        if (out != null) {
            out.close();
        }
    }

}