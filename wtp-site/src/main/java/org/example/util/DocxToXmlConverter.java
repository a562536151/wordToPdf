package org.example.util;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

public class DocxToXmlConverter{

    public void docxToXml() throws IOException {
        // docx文件路径
        String docxFilePath = "C:\\Users\\1\\Desktop\\workTppdf\\template.docx";
        // 解压目录
        String unzipPath = "C:\\Users\\1\\Desktop\\workTppdf";
        // 输出文件路径
        String xmlFilePath = "C:\\Users\\1\\Desktop\\workTppdf\\document.xml";

        // 创建ZipFile对象
        ZipFile zipFile = new ZipFile(docxFilePath);
        // 获取word/document.xml文件的ZipEntry对象
        ZipEntry entry = zipFile.getEntry("word/document.xml");
        // 通过ZipEntry对象获取输入流
        InputStream inputStream = zipFile.getInputStream(entry);
        // 创建输出流
        FileOutputStream outputStream = new FileOutputStream(xmlFilePath);
        // 缓冲区大小
        byte[] buffer = new byte[1024];
        int length;
        // 将输入流读入缓冲区，然后写入输出流
        while ((length = inputStream.read(buffer)) > 0) {
            outputStream.write(buffer, 0, length);
        }
        // 关闭输入输出流
        inputStream.close();
        outputStream.close();
        // 关闭ZipFile对象
        zipFile.close();

        System.out.println("生成文件：" + xmlFilePath);
    }
}
