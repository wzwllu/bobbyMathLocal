package com.wzw.bobby.servce;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

import javax.print.*;
import javax.print.attribute.DocAttributeSet;
import javax.print.attribute.HashDocAttributeSet;
import javax.print.attribute.HashPrintRequestAttributeSet;
import java.io.File;
import java.io.FileInputStream;

/**
 * @ Author     ：wuzhengwei.
 * @ Date       ：Created in 13:48 2020/10/14
 * @ Description：
 * @ Modified By：
 * @Version: $
 */
public class printService {

    public void print(){
        try {
        HashPrintRequestAttributeSet pras = new HashPrintRequestAttributeSet();
        // 设置打印格式，因为未确定类型，所以选择autosense
        DocFlavor flavor = DocFlavor.INPUT_STREAM.AUTOSENSE;
        System.out.println("打印文件类型为：==================="+flavor);
        //pras.add(MediaName.ISO_A4_TRANSPARENT);//A4纸张


        PrintService defaultService = PrintServiceLookup.lookupDefaultPrintService();
        System.out.println("打印工具选择打印机为：==================="+defaultService);

            DocPrintJob job = defaultService.createPrintJob(); // 创建打印作业
            FileInputStream fis = new FileInputStream(new File("D:\\1.txt")); // 构造待打印的文件流

            DocAttributeSet das = new HashDocAttributeSet();
            Doc doc = new SimpleDoc(fis, flavor, das);
            job.print(doc, pras);
        } catch (Exception e) {
            e.printStackTrace();
        }

    }



    public void printNew() {
        String path="D:\\bobby.docx";
        System.out.println("开始打印");
        ComThread.InitSTA();
        ActiveXComponent word=new ActiveXComponent("Word.Application");
        Dispatch doc=null;
        Dispatch.put(word, "Visible", new Variant(false));
        Dispatch docs=word.getProperty("Documents").toDispatch();
        doc=Dispatch.call(docs, "Open", path).toDispatch();
        System.out.println("开始");
        try {
            Dispatch.call(doc, "PrintOut");//打印
        } catch (Exception e) {
            e.printStackTrace();
            System.out.println("打印失败");
        }finally{
            try {
                if(doc!=null){
                    Dispatch.call(doc, "Close",new Variant(0));
                }
            } catch (Exception e2) {
                e2.printStackTrace();
            }
            //释放资源
            ComThread.Release();
        }
    }
}
