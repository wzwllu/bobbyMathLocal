package com.wzw.bobby.servce;

import com.wzw.bobby.bean.MathList;
import com.wzw.bobby.bean.MathShow;
import org.apache.poi.xwpf.converter.pdf.PdfConverter;
import org.apache.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.io.*;
import java.math.BigInteger;
import java.util.List;

/**
 * @ Author     ：wuzhengwei.
 * @ Date       ：Created in 17:12 2020/10/13
 * @ Description：
 * @ Modified By：
 * @Version: $
 */
public class docService {

    private List<MathList> list ;

    private XWPFDocument document= new XWPFDocument();

    public docService(List<MathList> list){
        this.list= list;
        this.createDoc();
    }


    public      XWPFDocument getDocument(){
        return document;
    }

    public  ByteArrayInputStream getInputStream() throws IOException {
        XWPFDocument document = new XWPFDocument ();//新建文档 后面NEW方法可以忽略
        ByteArrayOutputStream baos = new ByteArrayOutputStream();//二进制OutputStream
        document.write(baos);//文档写入流
        ByteArrayInputStream in = new ByteArrayInputStream(baos.toByteArray());//OutputStream写入InputStream二进制流
        return in;
    }


    public void addList(List<MathList> list){
       this.list=list;
        XWPFParagraph p = document.createParagraph();

        p.setPageBreak(true);
        this.createDoc();

    }

    public void createDoc(){
            //添加标题
            XWPFParagraph titleParagraph = document.createParagraph();
            //设置段落居中
            titleParagraph.setAlignment(ParagraphAlignment.CENTER);

            XWPFRun titleParagraphRun = titleParagraph.createRun();

            titleParagraphRun.setText("数 学 练 习");
            titleParagraphRun.setColor("000000");
            titleParagraphRun.setFontSize(30);

            //段落
            XWPFParagraph firstParagraph = document.createParagraph();
            firstParagraph.setAlignment(ParagraphAlignment.RIGHT);
            XWPFRun run = firstParagraph.createRun();
            run.setText("得分：        用时： 3     4       5       6       7       8       9       10     ");
            run.setColor("000000");
            run.setFontSize(12);

//        //设置段落背景颜色
//        CTShd cTShd = run.getCTR().addNewRPr().addNewShd();
//        cTShd.setVal(STShd.CLEAR);
//        cTShd.setFill("97FFFF");

            //换行
            XWPFParagraph paragraph1 = document.createParagraph();
            XWPFRun paragraphRun1 = paragraph1.createRun();
            paragraphRun1.setText("\r");

            //基本信息表格
            XWPFTable infoTable = document.createTable();
            //去表格边框
            infoTable.getCTTbl().getTblPr().unsetTblBorders();

//        infoTable.setInsideVBorder(XWPFTable.XWPFBorderType.NONE,0,0,"");



            //列宽自动分割
            CTTblWidth infoTableWidth = infoTable.getCTTbl().addNewTblPr().addNewTblW();
            infoTableWidth.setType(STTblWidth.DXA);
            infoTableWidth.setW(BigInteger.valueOf(9072));

            XWPFTableRow infoTableRowTwo = null;
            for(int i =0;i<list.size();i++){
                XWPFTableRow infoTableRowOne = infoTable.getRow(0);
                CTTrPr trPr = infoTableRowOne.getCtRow().addNewTrPr();

                CTHeight ht = trPr.addNewTrHeight();
                ht.setVal(BigInteger.valueOf(300));

                if(i==0){
                    //表格第一行

                    XWPFTableCell c0 =infoTableRowOne.getCell(0);
                    XWPFParagraph paragraph = c0.getParagraphs().get(0);
                    XWPFRun run0 = paragraph.createRun();
                    run0.setText(list.get(i).toNormalStr());
                    run0.setColor("000000");
                    run0.setFontSize(15);

//                c0.setText(list.get(i).toNormalStr());


                }else if(i==1||i==2||i==3){
                    XWPFTableCell cadd= infoTableRowOne.addNewTableCell();
                    XWPFParagraph paragraph = cadd.getParagraphs().get(0);
                    XWPFRun run0 = paragraph.createRun();
                    run0.setText(list.get(i).toNormalStr());
                    run0.setColor("000000");
                    run0.setFontSize(15);
//                .setText(list.get(i).toNormalStr());
                }else{
                    if(i%4==0){
                        infoTableRowTwo= infoTable.createRow();
                        CTTrPr trPr2 = infoTableRowTwo.getCtRow().addNewTrPr();
                        CTHeight ht2 = trPr2.addNewTrHeight();
                        ht2.setVal(BigInteger.valueOf(300));
                    }
                    XWPFTableCell cold = infoTableRowTwo.getCell(i%4);
                    XWPFParagraph paragraph = cold.getParagraphs().get(0);
                    XWPFRun run0 = paragraph.createRun();
                    run0.setText(list.get(i).toNormalStr());
                    run0.setColor("000000");
                    run0.setFontSize(15);

//                .setText(list.get(i).toNormalStr());
                }
            }







    }

    private void addHeader(){
        CTSectPr sectPr = document.getDocument().getBody().addNewSectPr();
        XWPFHeaderFooterPolicy policy = new XWPFHeaderFooterPolicy(document, sectPr);
        //添加页眉
        CTP ctpHeader = CTP.Factory.newInstance();
        CTR ctrHeader = ctpHeader.addNewR();
        CTText ctHeader = ctrHeader.addNewT();
        String headerText = "日期：_____年_____月_____日";
        ctHeader.setStringValue(headerText);
        XWPFParagraph headerParagraph = new XWPFParagraph(ctpHeader, document);
        //设置为右对齐
        headerParagraph.setAlignment(ParagraphAlignment.RIGHT);
        XWPFParagraph[] parsHeader = new XWPFParagraph[1];
        parsHeader[0] = headerParagraph;
        policy.createHeader(XWPFHeaderFooterPolicy.DEFAULT, parsHeader);

        //添加页脚
        CTP ctpFooter = CTP.Factory.newInstance();
        CTR ctrFooter = ctpFooter.addNewR();
        CTText ctFooter = ctrFooter.addNewT();
        String footerText = "上海市福山证大外国语小学  二年四班  吴轩宇";
        ctFooter.setStringValue(footerText);
        XWPFParagraph footerParagraph = new XWPFParagraph(ctpFooter, document);
        headerParagraph.setAlignment(ParagraphAlignment.CENTER);
        XWPFParagraph[] parsFooter = new XWPFParagraph[1];
        parsFooter[0] = footerParagraph;
        policy.createFooter(XWPFHeaderFooterPolicy.DEFAULT, parsFooter);
    }



    public void write2Docx()throws Exception {
        //Write the Document in file system
        this.addHeader();
//        this.createDoc();
//        PdfOptions options = PdfOptions.create();
//        OutputStream out = new FileOutputStream("D:\\create_table.pdf");
//        PdfConverter.getInstance().convert(document, out, options);
//
//        out.close();
        FileOutputStream out = new FileOutputStream(new File("D:\\bobby.docx"));
        document.write(out);
        out.close();


    }
//    public static void main(String[] args) throws Exception {
//        new docService().write2Docx();
//    }

}
