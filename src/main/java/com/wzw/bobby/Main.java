package com.wzw.bobby;

import com.wzw.bobby.bean.MathList;
import com.wzw.bobby.servce.addService;
import com.wzw.bobby.servce.docService;
import com.wzw.bobby.servce.fixService;
import com.wzw.bobby.servce.printService;

import java.util.Collections;
import java.util.List;

/**
 * @ Author     ：wuzhengwei.
 * @ Date       ：Created in 16:18 2020/10/13
 * @ Description：
 * @ Modified By：
 * @Version: $
 */
public class Main {

    public static void main(String[] args) throws Exception {

     int j =99;

        //生成word。
        docService ds =new docService(makelist());
        for(int i=0;i<j;i++) {
            ds.addList(makelist());
            System.out.println(i);
        }
//        ds.addList(makelist());

        ds.write2Docx();
        //打印
//         new printService().printNew();


    }


    private static  List<MathList> makelist(){
        String count = "100";
        String addScale = "50";
        String subScale = "50";
        String mutliScale = "0";
        String divScale = "0";
        String hard = "100";
        String addmax = "100";
        String submax = "100";
        String mutlimax = "10";
        String divmax = "10";



        //生产题目
        List<MathList> list = new fixService(Integer.valueOf(count),
                Integer.valueOf(addScale),
                Integer.valueOf(subScale),
                Integer.valueOf(mutliScale),
                Integer.valueOf(divScale),
                Integer.valueOf(hard),
                Integer.valueOf(addmax),
                Integer.valueOf(submax),
                Integer.valueOf(mutlimax),
                Integer.valueOf(divmax)).makeList();
        Collections.shuffle(list);
        return list;

    }
}
