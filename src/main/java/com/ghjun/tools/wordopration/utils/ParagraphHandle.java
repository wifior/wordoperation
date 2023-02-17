package com.ghjun.tools.wordopration.utils;

import lombok.extern.slf4j.Slf4j;

import java.io.IOException;

@Slf4j
public class ParagraphHandle {

    // 该正则用来匹配一个大题（例如：二、多选题）
    private static String regex = "([一|二|三|四|五|六|七|八|九|十]{1,3})([、.]{1})([\\u4E00-\\u9FA5\\s]+题)";

    public static void main(String[] args) throws IOException {
        wordParser();
    }/**
     * 根据poi将word文档逐级解析成实体对象
     *
     * @throws IOException
     */
    public static List<Question> wordParser() throws IOException {

        ArrayList<Question> list = new ArrayList<Question>();
        //第二种方式使用poi行读的方式逐步解析word文档拼接成实体对象
        XWPFDocument doc = new XWPFDocument(POIXMLDocument.openPackage("C:\\Users\\ywb\\Desktop\\试题分类.docx"));
//        List<XWPFParagraph> paragraphs = doc.getParagraphs();
        List<IBodyElement> paragraphs = doc.getBodyElements();

        System.out.println(paragraphs.size());
        //第一段固定为试卷名称
//        String title = ((XWPFParagraph) paragraphs.get(0)).getParagraphText();
        //第二段固定为什么类型的试卷
//        String subject = ((XWPFParagraph) paragraphs.get(1)).getParagraphText();

        for (int i = 2; i < paragraphs.size(); i++) {
            String paragraphText = ((XWPFParagraph) paragraphs.get(i)).getParagraphText();
            if (paragraphText.contains(QuestionTypeEnum.SINGLE_CHOICE_QUESTION.getType())) {
                int num = getChoiceQuestionList(i, QuestionTypeEnum.SINGLE_CHOICE_QUESTION.getType(), paragraphs, list);
                i = num;
            } else if (paragraphText.contains(QuestionTypeEnum.MULTIPLE_CHOICE.getType())) {
                int num = getChoiceQuestionList(i, QuestionTypeEnum.MULTIPLE_CHOICE.getType(), paragraphs, list);
                i = num;
            } else if (paragraphText.contains(QuestionTypeEnum.GAP_FILLING.getType())) {
                int num = getNotChoiceQuestionList(i, QuestionTypeEnum.GAP_FILLING.getType(), paragraphs, list);
                i = num;
            } else if (paragraphText.contains(QuestionTypeEnum.SUBJECTIVE_ITEM.getType())) {
                int num = getNotChoiceQuestionList(i, QuestionTypeEnum.SUBJECTIVE_ITEM.getType(), paragraphs, list);
                i = num;
            }
        }
        System.out.println(list);
        return list;
    }


    /**
     * 选择题的解析方法
     *
     * @param i
     * @param paragraphs
     * @return
     * @throws IOException
     */
    public static int getChoiceQuestionList(int i, String questionType, List<IBodyElement> paragraphs, List<Question> list) throws IOException {
        //该正则用来匹配一个大题（例如：二、多选题）
        Pattern compile = Pattern.compile(regex);
        i++;
        for (i = i; i < paragraphs.size(); i++) {
            String question = getContent(paragraphs.get(i), new StringBuilder());
            // 表示一种题型的结束
            if (compile.matcher(question).find()) {
                i--;
                break;
            }
            //获取word试卷题目中的题干信息
            if (question.contains(QuestionTypeEnum.TOPIC.getType())) {
                i++;
                for (i = i; i < paragraphs.size(); i++) {
                    String content = getContent(paragraphs.get(i), new StringBuilder());
                    if (content.contains(QuestionTypeEnum.CHOICE.getType())) {
                        break;
                    }
                    question += content;
                }
            }
            //获取题目中的选项信息
            String option = getContent(paragraphs.get(i), new StringBuilder());
            if (option.contains(QuestionTypeEnum.CHOICE.getType())) {
                i++;
                //如果标签不是以【答案】结尾，该题目为多行一题
                while (!((XWPFParagraph) paragraphs.get(i)).getParagraphText().contains(QuestionTypeEnum.ANSWER.getType())) {
                    String content = getContent(paragraphs.get(i), new StringBuilder());

                    option += content;
                    i++;
                }
            }
            //获取题目的答案信息
            String answer = getContent(paragraphs.get(i), new StringBuilder());
            if (answer.contains(QuestionTypeEnum.ANSWER.getType())) {
                i++;
                while (!((XWPFParagraph) paragraphs.get(i)).getParagraphText().contains(QuestionTypeEnum.ANALYSIS.getType())) {
                    String content = getContent(paragraphs.get(i), new StringBuilder());
                    answer += content;
                    i++;
                }
            }
            //获取题目的解析信息
            String analysis = getContent(paragraphs.get(i), new StringBuilder());
            if (analysis.contains(QuestionTypeEnum.ANALYSIS.getType())) {
                i++;
                while (!((XWPFParagraph) paragraphs.get(i)).getParagraphText().contains(QuestionTypeEnum.FINISH.getType())) {
                    String content = getContent(paragraphs.get(i), new StringBuilder());
                    analysis += content;
                    i++;
                }
            }

            if (StringUtils.isNotBlank(question) && StringUtils.isNotBlank(analysis) && StringUtils.isNotBlank(answer) && StringUtils.isNotBlank(option)) {
                Question q = new Question();
                //第一段固定为试卷名称
                String title = ((XWPFParagraph) paragraphs.get(0)).getParagraphText();
                //第二段固定为什么类型的试卷
                String subject = ((XWPFParagraph) paragraphs.get(1)).getParagraphText();
                q.setTitle(title);
                q.setSubject(subject);
                q.setQuestionType(questionType);
                q.setQuestion(question);
                q.setAnalysis(analysis);
                q.setRightAnswer(answer);
                q.setOption(option);
                Options options = divideOption(option);
                q.setOptions(options);
                list.add(q);
            }
        }
        System.out.println(i);
        return i;
    }
    /**
     * 将选项分隔出来
     *
     * @param option
     * @return
     */
    public static Options divideOption(String option) {
        //使用subString将答案分割
        String regex = "[A|B|C|D]、";
        Options options = new Options();
        String optionA = option.substring(option.indexOf("A"), option.indexOf("B"));
        String optionB = option.substring(option.indexOf("B"), option.indexOf("C"));
        String optionC = option.substring(option.indexOf("C"), option.indexOf("D"));
        String optionD = option.substring(option.indexOf("D"));
        options.setAnswerA(optionA);
        options.setAnswerB(optionB);
        options.setAnswerC(optionC);
        options.setAnswerD(optionD);
        return options;
    }/**
     *  段落处理器
     * @param content 拼接word内容
     * @param body word段落
     * @param imageParser 保存图片文件到磁盘
     */
    public static void handleParagraph(StringBuilder content, IBodyElement body, ImageParse imageParser) {
        log.info(">> 开始解析段落");
        XWPFParagraph p = (XWPFParagraph) body;
        if (p.isEmpty() || p.isWordWrap() || p.isPageBreak()) {
            return;
        }
//        String tagName = "p";
//        content.append("<" + tagName + ">");
      /*XWPFParagraph 有两个方法可以分别提出XWPFRun和CTOMath，但是不知道位置
      ParagraphChildOrderManager这个类是专门解决这个问题的
      */
        ParagraphChildOrderManager runOrMaths = new ParagraphChildOrderManager(p);
        List<Object> childList = runOrMaths.getChildList();

        for (Object child : childList) {
            if (child instanceof XWPFRun) {
                //处理段落中的文本以及图片
                handleParagraphRun(content, (XWPFRun) child, imageParser);
            } else if (child instanceof CTOMath) {
                // 处理word中存在的公式成mathML格式
                handleParagraphOMath(content, (CTOMath) child);
            } else if (child instanceof CTOMathPara) {
                //处理word中存在的公式
                handleParagraphOMath(content, (CTOMathPara) child);
            }
        }
//        content.append("</" + tagName + ">");
    }


    /**
     * 段落文本处理器
     *
     * @param content
     * @param run
     * @param imageParser
     */
    private static void handleParagraphRun(StringBuilder content, XWPFRun run, ImageParse imageParser) {
        // 有内嵌的图片
        List<XWPFPicture> pics = run.getEmbeddedPictures();
        if (pics != null && pics.size() > 0) {
            handleParagraphRunsImage(content, pics, imageParser);
        } else {
            //纯文本直接获取
            content.append(run.toString());
        }
    }

    /**
     * 处理图片
     *
     * @param content
     * @param pics
     * @param imageParser
     */
    private static void handleParagraphRunsImage(StringBuilder content, List<XWPFPicture> pics, ImageParse imageParser) {
        log.info(">> 开始解析段落中存在的图片");
        for (XWPFPicture pic : pics) {
            log.info(pic.getDescription());
            String path = imageParser.parse(pic.getPictureData().getData(),
                    pic.getPictureData().getFileName());
            log.debug("pic.getPictureData().getFileName()===" + pic.getPictureData().getFileName());

            CTPicture ctPicture = pic.getCTPicture();
            Node domNode = ctPicture.getDomNode();

            Node extNode = W3cNodeUtil.getChildChainNode(domNode, "pic:spPr", "a:ext");
            NamedNodeMap attributes = extNode.getAttributes();
            if (attributes != null && attributes.getNamedItem("cx") != null) {
                int width = WordMyUnits.emuToPx(new Double(attributes.getNamedItem("cx").getNodeValue()));
                int height = WordMyUnits.emuToPx(new Double(attributes.getNamedItem("cy").getNodeValue()));
                content.append(String.format("<img src=\"%s\" width=\"%d\" height=\"%d\" />", path, width, height));
            } else {
                content.append(String.format("<img src=\"%s\" />", path));
            }
        }
    }


    /**
     * 处理公式
     *
     * @param content
     * @param child
     */
    private static void handleParagraphOMath(StringBuilder content, CTOMath child) {
        log.info(">> 开始解析段落中的CTOMath公式");
        //将word中的omml格式转换成mathML
        String mathMLFromNode = OmmlUtils.getMathMLFromNode(child.xmlText());
        content.append(mathMLFromNode);
    }


    /**
     * 处理公式(将word中的omML格式的公式转换成mathML格式的工单)
     *
     * @param content
     * @param child
     */
    private static void handleParagraphOMath(StringBuilder content, CTOMathPara child) {
        log.info(">> 开始解析段落中的CTOMathPara公式");
        //将word中的omml格式转换成mathML
        String mathMLFromNode = OmmlUtils.getMathMLFromNode(child.xmlText());
        content.append(mathMLFromNode);
    }

}
