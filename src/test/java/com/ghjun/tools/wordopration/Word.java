package com.ghjun.tools.wordopration;

import org.apache.poi.xwpf.usermodel.*;

import java.io.FileInputStream;
import java.io.InputStream;
import java.util.List;

public class Word {

    // 读取docx文件
    public static void main(String[] args) {
        try {
            readDocx();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    public static void readDocx() throws Exception {
        InputStream input = new FileInputStream("F:\\ab.docx");
        // 定义document对象
        XWPFDocument xwpf = new XWPFDocument(input);

        // 获取document对象的子元素
        List<IBodyElement> ibs = xwpf.getBodyElements();
        for (IBodyElement ib : ibs) {
            // 对每个子元素的类型进行判断
            BodyElementType bet = ib.getElementType();
            // 如果是表格
            if (bet == BodyElementType.TABLE) {
                System.out.println("表格: " + ib.getPart());
            } else {
                // 段落
                XWPFParagraph paragraph = (XWPFParagraph) ib;
               // System.out.println("本段缩进: " + paragraph.getFirstLineIndent());

                List<XWPFRun> res = paragraph.getRuns();
                if (res.size() <= 0) {
                    System.out.println("空内容");
                }

                // 遍历每个run
                for (XWPFRun re : res) {
                    // 文本为空证明有图片在段落里
                    if ( re.text() == null || re.text().length() <= 0) {
                        // 有图片
                        if (re.getEmbeddedPictures().size() > 0) {
                            System.out.println("图片: " + re.getEmbeddedPictures());
                        } else {
                            // 公式域
                            if (re.getCTR().xmlText().indexOf("instrText") > 0) {
                                System.out.println("公式域");
                            } else {
                                System.out.println("其他内容");
                            }

                        }

                    } else {
                        System.out.println(new String("目标docx文件内容:".getBytes(), "UTF-8")
                                +re.text());
                    }

                }
            }

        }

        input.close();
    }

}
