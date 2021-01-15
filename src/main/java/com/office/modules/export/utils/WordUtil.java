package com.office.modules.export.utils;/**
 * Created by wei on 2020/12/25 9:41
 */

import com.office.modules.api.model.entity.ApiList;
import io.swagger.annotations.Api;
import java.io.FileOutputStream;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Set;
import java.util.stream.Collectors;
import lombok.ToString;
import org.apache.commons.lang.StringUtils;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.Borders;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.TextAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFStyle;
import org.apache.poi.xwpf.usermodel.XWPFStyles;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.tomcat.util.http.parser.TokenList;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDecimalNumber;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTFonts;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTOnOff;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTString;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTStyle;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTabStop;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STHdrFtr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STStyleType;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTabJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblWidth;

/**
 * @Author wei
 * @Date 2020/12/25 9:41
 **/
public class WordUtil {

    public static void main(String[] args) throws Exception {

    }

    public static LinkedHashMap<String, Object> getLinkedHashMap(List<ApiList> lists,String[] columns) throws Exception {
        LinkedHashMap<String, Object> params = new LinkedHashMap<String, Object>();
        // PUT页眉
        LinkedHashMap<String, Object> headerDoc = new LinkedHashMap<String, Object>();
        headerDoc.put("text", "");
        headerDoc.put("fontFamily", "新宋体");
        headerDoc.put("fontSize", 10);
        params.put("header", headerDoc);
        // PUT 正文
        params.put("body", saveBody(lists,columns));
        // PUT页脚
        LinkedHashMap<String, Object> footerDoc = new LinkedHashMap<String, Object>();
        footerDoc.put("textLeft", "");
        footerDoc.put("leftFontFamily", "新宋体");
        footerDoc.put("leftFontSize", 10);
        footerDoc.put("textRight", "智能打标接口调用清单");
        params.put("footer", footerDoc);
        return params;
    }

    public static List<LinkedHashMap<String, Object>> saveBody(List<ApiList> lists,String[] columns) throws Exception{
        List<LinkedHashMap<String, Object>> bodyDocs = new ArrayList<LinkedHashMap<String, Object>>();
        String sets= lists.stream().collect(Collectors.groupingBy(ApiList::getApiParentModel)).entrySet().stream().map(e->e.getKey()).collect(
                Collectors.joining(",","",""));
       // lists.stream().filter(l->l.getApiParentModel().contains(sets)).map(saveBody());
        // 1
        LinkedHashMap<String, Object> bodyDoc = new LinkedHashMap<String, Object>();
        bodyDoc.put("style", "标题 1");
        bodyDoc.put("text", "关键字管理");
        bodyDoc.put("fontFamily", "Courier");
        bodyDoc.put("fontSize", 16);
        // bodyDoc.put("textPosition", 20);
        List<LinkedHashMap<String, Object>> docs = new ArrayList<LinkedHashMap<String, Object>>();
        // 循环1-1
        LinkedHashMap<String, Object> docsDoc = new LinkedHashMap<String, Object>();
        docsDoc.put("text", "关键字新增接口");
        docsDoc.put("style", "标题 2");
        docsDoc.put("fontFamily", "宋体");
        docsDoc.put("fontSize", 12);
        docsDoc.put("tables",
                new String[][]{{"URL", "/api/v1/keywords/insert"}, {"请求方式", "POST"}, {"接口描述", "关键字新增接口"},
                        {"请求类型", "application/json"}, {"返回类型", "*/*"}, {"参数名", "数据类型", "参数类型", "是否必输", "说明"},
                        {"keyWordsName", "String", "body", "Y", ""}, {"状态码", "描述", "说明"}, {"200", "OK", "OK"},
                        {"201", "CREATED", "CREATED"}, {"401", "Unauthorized", "Unauthorized"},
                        {"403", "Forbidden", "Forbidden"}, {"404", "Not Found", "Not Found"},
                        {"示例", "{\"keyWordsName\": \"\"}"}, {"请求参数", "{\"keyWordsName\": \"\"}"}, {"返回值", ""}});
        docs.add(saveForDocument(docsDoc));
        // 循环1-2
        docsDoc = new LinkedHashMap<String, Object>();
        docsDoc.put("text", "关键字发布更新");
        docsDoc.put("style", "标题 2");
        docsDoc.put("fontFamily", "宋体");
        docsDoc.put("fontSize", 12);
        docsDoc.put("tables",
                new String[][]{{"URL", "/api/v1/keywords/publish/update"}, {"请求方式", "POST"},
                        {"接口描述", "关键字发布更新"}, {"请求类型", "application/json"}, {"返回类型", "*/*"},
                        {"参数名", "数据类型", "参数类型", "是否必输", "说明"}, {"keyWordsName", "String", "body", "Y", ""},
                        {"状态码", "描述", "说明"}, {"200", "OK", "OK"}, {"201", "CREATED", "CREATED"},
                        {"401", "Unauthorized", "Unauthorized"}, {"403", "Forbidden", "Forbidden"},
                        {"404", "Not Found", "Not Found"}, {"示例", "{\"keyWordsName\": \"\"}"},
                        {"请求参数", "{\"keyWordsName\": \"\"}"}, {"返回值", ""}});
        docs.add(saveForDocument(docsDoc));
        // 循环1-3
        docsDoc = new LinkedHashMap<String, Object>();
        docsDoc.put("text", "关键字取消发布更新");
        docsDoc.put("style", "标题 2");
        docsDoc.put("fontFamily", "宋体");
        docsDoc.put("fontSize", 12);
        docsDoc.put("tables",
                new String[][]{{"URL", "/api/v1/keywords/unpublish/update"}, {"请求方式", "POST"},
                        {"接口描述", "关键字取消发布更新"}, {"请求类型", "application/json"}, {"返回类型", "*/*"},
                        {"参数名", "数据类型", "参数类型", "是否必输", "说明"}, {"keyWordsName", "String", "body", "Y", ""},
                        {"状态码", "描述", "说明"}, {"200", "OK", "OK"}, {"201", "CREATED", "CREATED"},
                        {"401", "Unauthorized", "Unauthorized"}, {"403", "Forbidden", "Forbidden"},
                        {"404", "Not Found", "Not Found"}, {"示例", "{\"keyWordsName\": \"\"}"},
                        {"请求参数", "{\"keyWordsName\": \"\"}"}, {"返回值", ""}});
        docs.add(saveForDocument(docsDoc));
        // 循环1-4
        docsDoc = new LinkedHashMap<String, Object>();
        docsDoc.put("text", "关键字未发布查询");
        docsDoc.put("style", "标题 2");
        docsDoc.put("fontFamily", "宋体");
        docsDoc.put("fontSize", 12);
        docsDoc.put("tables",
                new String[][]{{"URL", "/api/v1/keywords/unpublish/list"}, {"请求方式", "GET"},
                        {"接口描述", "关键字未发布查询"}, {"请求类型", "application/json"}, {"返回类型", "*/*"},
                        {"参数名", "数据类型", "参数类型", "是否必输", "说明"}, {"keyWordsName", "String", "body", "N", ""},
                        {"状态码", "描述", "说明"}, {"200", "OK", "OK"}, {"201", "CREATED", "CREATED"},
                        {"401", "Unauthorized", "Unauthorized"}, {"403", "Forbidden", "Forbidden"},
                        {"404", "Not Found", "Not Found"}, {"示例", "{\"keyWordsName\": \"\"}"},
                        {"请求参数", "{\"keyWordsName\": \"\"}"}, {"返回值", ""}});
        docs.add(saveForDocument(docsDoc));
        // 循环1-5
        docsDoc = new LinkedHashMap<String, Object>();
        docsDoc.put("text", "关键字已发布查询");
        docsDoc.put("style", "标题 2");
        docsDoc.put("fontFamily", "宋体");
        docsDoc.put("fontSize", 12);
        docsDoc.put("tables",
                new String[][]{{"URL", "/api/v1/keywords/publish/list"}, {"请求方式", "GET"}, {"接口描述", "关键字已发布查询"},
                        {"请求类型", "application/json"}, {"返回类型", "*/*"}, {"参数名", "数据类型", "参数类型", "是否必输", "说明"},
                        {"keyWordsName", "String", "body", "N", ""}, {"状态码", "描述", "说明"}, {"200", "OK", "OK"},
                        {"201", "CREATED", "CREATED"}, {"401", "Unauthorized", "Unauthorized"},
                        {"403", "Forbidden", "Forbidden"}, {"404", "Not Found", "Not Found"},
                        {"示例", "{\"keyWordsName\": \"\"}"}, {"请求参数", "{\"keyWordsName\": \"\"}"}, {"返回值", ""}});
        docs.add(saveForDocument(docsDoc));
        // 循环1-6
        docsDoc = new LinkedHashMap<String, Object>();
        docsDoc.put("text", "关键字查询应用");
        docsDoc.put("style", "标题 2");
        docsDoc.put("fontFamily", "宋体");
        docsDoc.put("fontSize", 12);
        docsDoc.put("tables",
                new String[][]{{"URL", "/api/v1/keywords/application/list"}, {"请求方式", "GET"},
                        {"接口描述", "关键字查询应用"}, {"请求类型", "application/json"}, {"返回类型", "*/*"},
                        {"参数名", "数据类型", "参数类型", "是否必输", "说明"}, {"keyWordsName", "String", "body", "N", ""},
                        {"状态码", "描述", "说明"}, {"200", "OK", "OK"}, {"201", "CREATED", "CREATED"},
                        {"401", "Unauthorized", "Unauthorized"}, {"403", "Forbidden", "Forbidden"},
                        {"404", "Not Found", "Not Found"}, {"示例", "{\"keyWordsName\": \"\"}"},
                        {"请求参数", ""}, {"返回值", ""}});
        docs.add(saveForDocument(docsDoc));
        // 循环1-6
        docsDoc = new LinkedHashMap<String, Object>();
        docsDoc.put("text", "关键字删除");
        docsDoc.put("style", "标题 2");
        docsDoc.put("fontFamily", "宋体");
        docsDoc.put("fontSize", 12);
        docsDoc.put("tables",
                new String[][]{{"URL", "/api/v1/keywords/unpublish/delete"}, {"请求方式", "POST"},
                        {"接口描述", "关键字删除"}, {"请求类型", "application/json"}, {"返回类型", "*/*"},
                        {"参数名", "数据类型", "参数类型", "是否必输", "说明"}, {"keyWordsName", "String", "body", "Y", ""},
                        {"状态码", "描述", "说明"}, {"200", "OK", "OK"}, {"201", "CREATED", "CREATED"},
                        {"401", "Unauthorized", "Unauthorized"}, {"403", "Forbidden", "Forbidden"},
                        {"404", "Not Found", "Not Found"}, {"示例", "{\"keyWordsName\": \"\"}"},
                        {"请求参数", "{\"keyWordsName\": \"\"}"}, {"返回值", ""}});
        docs.add(saveForDocument(docsDoc));
        bodyDoc.put("docBody", docs);
        bodyDocs.add(bodyDoc);

        // 2
        bodyDoc = new LinkedHashMap<String, Object>();
        bodyDoc.put("style", "标题 1");
        bodyDoc.put("text", "标签审核管理");
        bodyDoc.put("fontFamily", "Courier");
        bodyDoc.put("fontSize", 16);
        // bodyDoc.put("textPosition", 20);
        docs = new ArrayList<LinkedHashMap<String, Object>>();
        // 循环2-1
        docsDoc = new LinkedHashMap<String, Object>();
        docsDoc.put("text", "查询待审核标签信息接口列表");
        docsDoc.put("style", "标题 2");
        docsDoc.put("fontFamily", "宋体");
        docsDoc.put("fontSize", 12);
        docsDoc.put("tables",
                new String[][]{{"URL", "/api/v1/label/unchecked/list"}, {"请求方式", "GET"},
                        {"接口描述", "查询待审核标签信息接口列表"}, {"请求类型", "application/json"}, {"返回类型", "*/*"},
                        {"参数名", "数据类型", "参数类型", "是否必输", "说明"}, {"keyWordsName", "String", "body", "N", ""},
                        {"labelListName", "String", "body", "N", ""}, {"creatorPartName", "String", "body", "N", ""},
                        {"状态码", "描述", "说明"}, {"200", "OK", "OK"}, {"201", "CREATED", "CREATED"},
                        {"401", "Unauthorized", "Unauthorized"}, {"403", "Forbidden", "Forbidden"},
                        {"404", "Not Found", "Not Found"},
                        {"示例", "{\"keyWordsName\":\"\",\"labelListName\":\"\",\"creatorPartName\":\"\"}"},
                        {"请求参数", ""}, {"返回值", ""}});
        docs.add(saveForDocument(docsDoc));
        // 循环2-2
        docsDoc = new LinkedHashMap<String, Object>();
        docsDoc.put("text", "查询已审核标签信息接口列表");
        docsDoc.put("style", "标题 2");
        docsDoc.put("fontFamily", "宋体");
        docsDoc.put("fontSize", 12);
        docsDoc.put("tables",
                new String[][]{{"URL", "/api/v1/label/checked/list"}, {"请求方式", "GET"},
                        {"接口描述", "查询已审核标签信息接口列表"}, {"请求类型", "application/json"}, {"返回类型", "*/*"},
                        {"参数名", "数据类型", "参数类型", "是否必输", "说明"}, {"keyWordsName", "String", "body", "N", " "},
                        {"labelListName", "String", "body", "N", " "}, {"creatorPartName", "String", "body", "N", " "},
                        {"状态码", "描述", "说明"}, {"200", "OK", "OK"}, {"201", "CREATED", "CREATED"},
                        {"401", "Unauthorized", "Unauthorized"}, {"403", "Forbidden", "Forbidden"},
                        {"404", "Not Found", "Not Found"},
                        {"示例", "{\"keyWordsName\":\"\",\"labelListName\":\"\",\"creatorPartName\":\"\"}"},
                        {"请求参数", ""}, {"返回值", ""}});
        docs.add(saveForDocument(docsDoc));
        // 循环2-3
        docsDoc = new LinkedHashMap<String, Object>();
        docsDoc.put("text", "查询单个标签信息接口");
        docsDoc.put("style", "标题 2");
        docsDoc.put("fontFamily", "宋体");
        docsDoc.put("fontSize", 12);
        docsDoc.put("tables",
                new String[][]{{"URL", "/api/v1/label/select"}, {"请求方式", "GET"}, {"接口描述", "查询单个标签信息接口"},
                        {"请求类型", "application/json"}, {"返回类型", "*/*"}, {"参数名", "数据类型", "参数类型", "是否必输", "说明"},
                        {"id", "String", "body", "Y", ""}, {"状态码", "描述", "说明"}, {"200", "OK", "OK"},
                        {"201", "CREATED", "CREATED"}, {"401", "Unauthorized", "Unauthorized"},
                        {"403", "Forbidden", "Forbidden"}, {"404", "Not Found", "Not Found"}, {"示例", "{\"id\":\"\"}"},
                        {"请求参数", ""}, {"返回值", ""}});
        docs.add(saveForDocument(docsDoc));
        // 循环2-4
        docsDoc = new LinkedHashMap<String, Object>();
        docsDoc.put("text", "标签审核操作信息接口");
        docsDoc.put("style", "标题 2");
        docsDoc.put("fontFamily", "宋体");
        docsDoc.put("fontSize", 12);
        docsDoc.put("tables",
                new String[][]{{"URL", "/api/v1/label/operate"}, {"请求方式", "POST"}, {"接口描述", "标签审核操作信息接口"},
                        {"请求类型", "application/json"}, {"返回类型", "*/*"}, {"参数名", "数据类型", "参数类型", "是否必输", "说明"},
                        {"id", "String", "body", "Y", ""}, {"type", "String", "body", "Y", ""}, {"状态码", "描述", "说明"},
                        {"200", "OK", "OK"},
                        {"201", "CREATED", "CREATED"}, {"401", "Unauthorized", "Unauthorized"},
                        {"403", "Forbidden", "Forbidden"}, {"404", "Not Found", "Not Found"},
                        {"示例", "{\"id\":\"\",\"type\":\"\"}"},
                        {"请求参数", ""}, {"返回值", ""}});
        docs.add(saveForDocument(docsDoc));
        bodyDoc.put("docBody", docs);
        bodyDocs.add(bodyDoc);

        return bodyDocs;
    }


    public static LinkedHashMap<String, Object> saveForDocument(LinkedHashMap<String, Object> params) throws Exception {
        LinkedHashMap<String, Object> linkedHashMap = new LinkedHashMap<String, Object>();
        linkedHashMap.put("text", params.get("text"));
        linkedHashMap.put("style", params.get("style"));
        linkedHashMap.put("fontFamily", params.get("fontFamily"));
        linkedHashMap.put("fontSize", params.getOrDefault("fontSize", 12));
        // tables
        List<List> tables = new ArrayList<>();
        String[][] values = (String[][]) params.get("tables");
        for (String[] value : values) {
            tables.add(Arrays.asList(value));
        }
        linkedHashMap.put("tables", tables);
        return linkedHashMap;
    }

    /**
     * 导出WORD方法 文档数组转换成WORD文件
     *
     * @return
     * @throws Exception
     */
    public static boolean docs2Word(LinkedHashMap<String, Object> param, String filePath) throws Exception {
        // 创建WORD空白文档
        XWPFDocument xwpfDoc = new XWPFDocument();
        List<LinkedHashMap<String, Object>> bodyLists = (List<LinkedHashMap<String, Object>>) param.get("body");
        int numId = 1;
        int k = 0;
        for (LinkedHashMap<String, Object> docBody : bodyLists) {
            createParagraph(xwpfDoc, numId, 1, docBody);
            int i = 1;
            List<LinkedHashMap<String, Object>> docbodyDocs = (List<LinkedHashMap<String, Object>>) docBody
                    .get("docBody");
            for (LinkedHashMap<String, Object> docbody : docbodyDocs) {
                createParagraph(xwpfDoc, i, 2, docbody);
                List<List> lists = (List<List>) docbody.get("tables");
                XWPFTable table = createTable(xwpfDoc, lists.size(), 5);
                int rows = 1;
                for (List<String> list : lists) {
                    createTalbeAndCell(xwpfDoc, k, rows, 5, list);
                    rows++;
                    System.out.println(numId + "--" + i + "--" + k + "--" + rows);
                }
                k++;

            }
            i++;
            numId++;
        }

        // 设置页眉
        /*
         * 对页眉段落作处理，使公司logo图片在页眉左边，公司全称在页眉右边
         */
        LinkedHashMap<String, Object> docHeader = (LinkedHashMap<String, Object>) param.get("header");
        CTSectPr sectPr = xwpfDoc.getDocument().getBody().addNewSectPr();
        XWPFHeaderFooterPolicy headerFooterPolicy = new XWPFHeaderFooterPolicy(xwpfDoc, sectPr);
        XWPFHeader header = headerFooterPolicy.createHeader(XWPFHeaderFooterPolicy.DEFAULT);
        List<XWPFParagraph> paragraphList = header.getParagraphs();
        for (XWPFParagraph paragraphHeader : paragraphList) {
            paragraphHeader.setAlignment(ParagraphAlignment.LEFT);
            paragraphHeader.setBorderBottom(Borders.THICK);
            CTTabStop tabStopHeader = paragraphHeader.getCTP().getPPr().addNewTabs().addNewTab();
            tabStopHeader.setVal(STTabJc.RIGHT);
            int twipsPerInchHeader = 1440;
            tabStopHeader.setPos(BigInteger.valueOf(6 * twipsPerInchHeader));
            XWPFRun runHeader = paragraphHeader.createRun();
            setXWPFRunStyle(runHeader, String.valueOf(docHeader.get("fontFamily")),
                    Integer.valueOf(String.valueOf(docHeader.getOrDefault("fontSize", 10))));
            runHeader.setText(String.valueOf(docHeader.get("text")));
        }

        // 设置页脚
        /*
         * 生成页脚段落 给段落设置宽度为占满一行 为公司地址和公司电话左对齐，页码右对齐创造条件
         */
        XWPFFooter footer = headerFooterPolicy.createFooter(STHdrFtr.DEFAULT);
        List<XWPFParagraph> paragraphs = footer.getParagraphs();
        for (XWPFParagraph paragraphFooter : paragraphs) {
            paragraphFooter.setAlignment(ParagraphAlignment.LEFT);
            paragraphFooter.setVerticalAlignment(TextAlignment.CENTER);
            paragraphFooter.setBorderTop(Borders.THICK);
            CTTabStop tabStopFooter = paragraphFooter.getCTP().getPPr().addNewTabs().addNewTab();
            tabStopFooter.setVal(STTabJc.RIGHT);
            int twipsPerInchFooter = 1440;
            tabStopFooter.setPos(BigInteger.valueOf(6 * twipsPerInchFooter));
            /*
             * 给段落创建元素 设置元素字面为公司地址+公司主页链接
             */
            LinkedHashMap<String, Object> docFooter = (LinkedHashMap<String, Object>) param.get("footer");
            XWPFRun runFooter = paragraphFooter.createRun();
            runFooter.setText(String.valueOf(docFooter.get("textLeft")));
            setXWPFRunStyle(runFooter, String.valueOf(docFooter.get("leftFontFamily")),
                    Integer.valueOf(String.valueOf(docFooter.getOrDefault("leftFontSize", 10))));
            runFooter.addTab();
            runFooter.setText(String.valueOf(docFooter.get("textRight")));
            runFooter.addTab();
        }
        // 输出文档
        FileOutputStream out = new FileOutputStream(filePath);
        xwpfDoc.write(out);
        out.close();
        return true;
    }

    /**
     * 导出WORD方法 文档转换成WORD文件
     *
     * @param xwpfDoc WORD文档
     *         文档数据
     * @return
     * @throws Exception
     */
    public static XWPFParagraph createParagraph(XWPFDocument xwpfDoc, int numId, int level,
            LinkedHashMap<String, Object> linkedHashMap)
            throws Exception {
        // 服务名标题
        XWPFParagraph p1 = xwpfDoc.createParagraph();
        addCustomHeadingStyle(xwpfDoc, String.valueOf(linkedHashMap.get("style")), level);
        p1.setStyle(String.valueOf(linkedHashMap.get("style")));
        p1.setNumID(BigInteger.valueOf(numId));
        p1.setNumILvl(BigInteger.valueOf(level));
        p1.setAlignment(ParagraphAlignment.LEFT);
        XWPFRun r1 = p1.createRun();
        r1.setBold(true);
        r1.setText(String.valueOf(linkedHashMap.get("text")));
        r1.setFontFamily(String.valueOf(linkedHashMap.get("fontFamily")));
        r1.setFontSize(Integer.valueOf(String.valueOf(linkedHashMap.getOrDefault("fontSize", 12))));
        r1.setTextPosition(Integer.valueOf(String.valueOf(linkedHashMap.getOrDefault("textPosition", 0))));
        return p1;
    }

    public static void createTalbeAndCell(XWPFDocument xwpfDoc, int tableNum, int rows, int cols, List<String> list) {
        XWPFTable table = xwpfDoc.getTables().get(tableNum);
        if (list.size() != 5) {
            mergeCellsHorizontal(table, rows - 1, list.size() - 1, 4);
        }
        int i = 0;
        for (String value : list) {
            setXWPFTableCellText(table.getRow(rows - 1).getCell(i), i, value);
            i++;
        }
    }

    /**
     * 创建表格
     *
     * @param xwpfDoc
     * @param rows
     * @param cols
     * @return
     */
    public static XWPFTable createTable(XWPFDocument xwpfDoc, int rows, int cols) {
        // 列宽自动分割
        XWPFTable table = xwpfDoc.createTable(rows, cols);
        CTTblWidth infoTableWidth = table.getCTTbl().addNewTblPr().addNewTblW();
        int tableWidth = 8690;

        if (cols == 2) {
            int cellWidth = tableWidth / 9;
            XWPFTableCell cell = table.getRow(rows - 1).getCell(0);
            CTTblWidth newTcW = cell.getCTTc().addNewTcPr().addNewTcW();
            newTcW.setType(STTblWidth.DXA);
            newTcW.setW(BigInteger.valueOf(cellWidth));
            XWPFTableCell cell2 = table.getRow(rows - 1).getCell(1);
            CTTblWidth newTcW2 = cell2.getCTTc().addNewTcPr().addNewTcW();
            newTcW2.setType(STTblWidth.DXA);
            newTcW2.setW(BigInteger.valueOf(cellWidth * 4));
        } else if (cols == 3) {
            int cellWidth = tableWidth / 9;
            XWPFTableCell cell = table.getRow(rows - 1).getCell(0);
            CTTblWidth newTcW = cell.getCTTc().addNewTcPr().addNewTcW();
            newTcW.setType(STTblWidth.DXA);
            newTcW.setW(BigInteger.valueOf(cellWidth));
            XWPFTableCell cell2 = table.getRow(rows - 1).getCell(1);
            CTTblWidth newTcW2 = cell2.getCTTc().addNewTcPr().addNewTcW();
            newTcW2.setType(STTblWidth.DXA);
            newTcW2.setW(BigInteger.valueOf(cellWidth));
            XWPFTableCell cell3 = table.getRow(rows - 1).getCell(2);
            CTTblWidth newTcW3 = cell3.getCTTc().addNewTcPr().addNewTcW();
            newTcW3.setType(STTblWidth.DXA);
            newTcW3.setW(BigInteger.valueOf(cellWidth * 3));
        } else if (cols == 5) {
            infoTableWidth.setType(STTblWidth.DXA);
            infoTableWidth.setW(BigInteger.valueOf(tableWidth));
        }
        return table;
    }

    /**
     * word跨列合并单元格
     *
     * @param table
     * @param row
     * @param fromCell
     * @param toCell
     */
    public static void mergeCellsHorizontal(XWPFTable table, int row, int fromCell, int toCell) {
        for (int cellIndex = fromCell; cellIndex <= toCell; cellIndex++) {
            XWPFTableCell cell = table.getRow(row).getCell(cellIndex);
            CTTblWidth newTcW = null;
            if (cellIndex == fromCell) {
                // The first merged cell is set with RESTART merge value
                cell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.RESTART);
                newTcW = cell.getCTTc().addNewTcPr().addNewTcW();
                newTcW.setType(STTblWidth.DXA);
                newTcW.setW(BigInteger.valueOf(9072 * 4));
            } else {
                // Cells which join (merge) the first one, are set with CONTINUE
                newTcW = cell.getCTTc().addNewTcPr().addNewTcW();
                newTcW.setType(STTblWidth.DXA);
                newTcW.setW(BigInteger.valueOf(9072));
                cell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.CONTINUE);
            }
        }
    }

    /**
     * word跨行并单元格
     *
     * @param table
     * @param col
     * @param fromRow
     * @param toRow
     */
    public static void mergeCellsVertically(XWPFTable table, int col, int fromRow, int toRow) {
        for (int rowIndex = fromRow; rowIndex <= toRow; rowIndex++) {
            XWPFTableCell cell = table.getRow(rowIndex).getCell(col);
            if (rowIndex == fromRow) {
                // The first merged cell is set with RESTART merge value
                cell.getCTTc().addNewTcPr().addNewVMerge().setVal(STMerge.RESTART);
            } else {
                // Cells which join (merge) the first one, are set with CONTINUE
                cell.getCTTc().addNewTcPr().addNewVMerge().setVal(STMerge.CONTINUE);
            }
        }
    }

    /**
     * 设置单元格的值
     *
     * @param cell
     * @param cellText
     */
    public static void setXWPFTableCellText(XWPFTableCell cell, int i, String cellText) {
        CTP ctp = CTP.Factory.newInstance();
        XWPFParagraph p = new XWPFParagraph(ctp, cell);
        if (i == 0) {
            p.setAlignment(switchParagraphAlignment("center"));
        } else {
            p.setAlignment(switchParagraphAlignment("left"));
        }
        XWPFRun run = p.createRun();
        run.setText(cellText);
        if (i == 0) {
            run.setBold(true);
        }
        CTRPr rpr = run.getCTR().isSetRPr() ? run.getCTR().getRPr() : run.getCTR().addNewRPr();
        CTFonts fonts = rpr.isSetRFonts() ? rpr.getRFonts() : rpr.addNewRFonts();
        fonts.setAscii("仿宋");
        fonts.setEastAsia("仿宋");
        fonts.setHAnsi("仿宋");
        cell.setParagraph(p);
    }

    /**
     * 单元格对齐位置
     *
     * @return
     */
    public static ParagraphAlignment switchParagraphAlignment(String position) {
        switch (position.toLowerCase()) {
            case "left":
                return ParagraphAlignment.LEFT;
            case "center":
                return ParagraphAlignment.CENTER;
            case "right":
                return ParagraphAlignment.RIGHT;
            default:
                return ParagraphAlignment.LEFT;
        }
    }

    // private void nodeTe() {
    // // TODO Auto-generated method stub
    // for (int len = 0; len < items.length; len += 2) {
    // // 2.服务名称
    // XWPFParagraph p2 = xwpfDoc.createParagraph();
    // p2.setStyle("2");
    // p2.setAlignment(ParagraphAlignment.LEFT);
    // XWPFRun r2 = p2.createRun();
    // r2.setBold(true);
    // r2.setText(i + "." + index + "." + items[len]);
    //
    // // 服务名称正文
    // if (items[len + 1].equalsIgnoreCase("{table}")) {
    // // 输入参数正文表格
    // // 设置表头行
    // // 表头数据
    // List<String> headerTextList = new ArrayList<String>();
    // headerTextList.add("序号");
    // headerTextList.add("参数ID");
    // headerTextList.add("数据类型");
    // headerTextList.add("是否必输");
    // headerTextList.add("中文名称");
    // headerTextList.add("示例值");
    // int rows = 1;
    // XWPFTable table = xwpfDoc.createTable(rows, headerTextList.size());
    // // 列宽自动分割
    // CTTblWidth infoTableWidth = table.getCTTbl().addNewTblPr().addNewTblW();
    // infoTableWidth.setType(STTblWidth.DXA);
    // infoTableWidth.setW(BigInteger.valueOf(9072));
    // // 设置表头值
    // XWPFTableRow rowHeader = table.getRow(0);
    // int t = 0;
    // for (String headerText : headerTextList) {
    // XWPFParagraph cellParagraphC =
    // rowHeader.getCell(t).getParagraphs().get(0);
    // cellParagraphC.setAlignment(ParagraphAlignment.CENTER); // 设置表格内容居中
    // XWPFRun xwpfRun = cellParagraphC.createRun();
    // xwpfRun.setText(headerText);
    // xwpfRun.setBold(true);
    // t++;
    // }
    // // 设置内容行
    // // 设置内容数据
    // List<Document> inParamsList = (List<Document>) JSONArray.parse((String)
    // doc.get("inParams"));
    //
    // // 设置内容数据列名
    // List<String> headerTextColumnList = new ArrayList<String>();
    // headerTextColumnList.add("id");
    // headerTextColumnList.add("fieldId");
    // headerTextColumnList.add("fieldType");
    // headerTextColumnList.add("required");
    // headerTextColumnList.add("fieldName");
    // headerTextColumnList.add("sampleValue");
    // // 设置内容值
    // int k1 = 1;
    // if (inParamsList == null || inParamsList.size() == 0) {
    // table.createRow();
    // } else {
    // for (Document inParamsDoc : inParamsList) {
    // XWPFTableRow rowsContent = table.createRow();
    // for (int k = 0; k < headerTextColumnList.size(); k++) {
    // XWPFParagraph cellParagraphC =
    // rowsContent.getCell(k).getParagraphs().get(0);
    // cellParagraphC.setAlignment(ParagraphAlignment.CENTER); // 设置表格内容居中
    // XWPFRun cellParagraphRunC = cellParagraphC.createRun();
    // cellParagraphRunC.setFontSize(10); // 设置表格内容字号
    // String value = "";
    // if (headerTextColumnList.get(k).equalsIgnoreCase("id")) {
    // value = "" + (k1);
    // } else if (headerTextColumnList.get(k).equalsIgnoreCase("required")) {
    // boolean required = inParamsDoc.getBoolean(headerTextColumnList.get(k));
    // value = required ? "是" : "否";
    // } else {
    // value = inParamsDoc.getString(headerTextColumnList.get(k));
    // }
    // cellParagraphRunC.setText(value); // 单元格段落加载内容
    // }
    // k1++;
    // }
    // }
    // } else {
    // XWPFParagraph p3 = xwpfDoc.createParagraph();
    // p3.setAlignment(ParagraphAlignment.LEFT);
    // XWPFRun r3 = p3.createRun();
    // if (items[len + 1].toString().indexOf("|") != -1) {
    // String[] itemArr = items[len + 1].toString().split("\\|");
    // for (String item : itemArr) {
    // r3.setText(item);
    // r3.addCarriageReturn();
    // }
    // } else {
    // r3.addTab();// 添加TAB缩进
    // r3.setText(items[len + 1]);
    // }
    // }
    // index++;
    // }
    // }

    /**
     * 增加自定义标题样式。这里用的是stackoverflow的源码
     *
     * @param docxDocument 目标文档
     * @param strStyleId 样式名称
     * @param headingLevel 样式级别
     */
    public static void addCustomHeadingStyle(XWPFDocument docxDocument, String strStyleId, int headingLevel) {
        CTStyle ctStyle = CTStyle.Factory.newInstance();
        ctStyle.setStyleId(strStyleId);
        CTString styleName = CTString.Factory.newInstance();
        styleName.setVal(strStyleId);
        ctStyle.setName(styleName);
        CTDecimalNumber indentNumber = CTDecimalNumber.Factory.newInstance();
        indentNumber.setVal(BigInteger.valueOf(headingLevel));
        // lower number > style is more prominent in the formats bar
        ctStyle.setUiPriority(indentNumber);
        CTOnOff onoffnull = CTOnOff.Factory.newInstance();
        ctStyle.setUnhideWhenUsed(onoffnull);
        // style shows up in the formats bar
        ctStyle.setQFormat(onoffnull);
        // style defines a heading of the given level
        CTPPr ppr = CTPPr.Factory.newInstance();
        ppr.setOutlineLvl(indentNumber);
        ctStyle.setPPr(ppr);
        XWPFStyle style = new XWPFStyle(ctStyle);
        // is a null op if already defined
        XWPFStyles styles = docxDocument.createStyles();
        style.setType(STStyleType.PARAGRAPH);
        styles.addStyle(style);
    }

    /**
     * 设置页脚的字体样式
     *
     * @param r1 段落元素
     */
    public static void setXWPFRunStyle(XWPFRun r1, String font, int fontSize) {
        r1.setFontSize(fontSize);
        CTRPr rpr = r1.getCTR().isSetRPr() ? r1.getCTR().getRPr() : r1.getCTR().addNewRPr();
        CTFonts fonts = rpr.isSetRFonts() ? rpr.getRFonts() : rpr.addNewRFonts();
        fonts.setAscii(font);
        fonts.setEastAsia(font);
        fonts.setHAnsi(font);
    }

    /**
     * Word中的大纲级别，可以通过getPPr().getOutlineLvl()直接提取，但需要注意，Word中段落级别，通过如下三种方式定义：
     * 1、直接对段落进行定义； 2、对段落的样式进行定义； 3、对段落样式的基础样式进行定义。
     * 因此，在通过“getPPr().getOutlineLvl()”提取时，需要依次在如上三处读取。
     *
     * @param doc
     * @param para
     * @return
     */
    public static String getTitleLvl(XWPFDocument doc, XWPFParagraph para) {
        String titleLvl = "";
        try {
            // 判断该段落是否设置了大纲级别
            if (para.getCTP().getPPr().getOutlineLvl() != null) {
                // System.out.println("getCTP()");
                // System.out.println(para.getParagraphText());
                // System.out.println(para.getCTP().getPPr().getOutlineLvl().getVal());

                return String.valueOf(para.getCTP().getPPr().getOutlineLvl().getVal());
            }
        } catch (Exception e) {

        }

        try {
            // 判断该段落的样式是否设置了大纲级别
            if (doc.getStyles().getStyle(para.getStyle()).getCTStyle().getPPr().getOutlineLvl() != null) {

                // System.out.println("getStyle");
                // System.out.println(para.getParagraphText());
                // System.out.println(doc.getStyles().getStyle(para.getStyle()).getCTStyle().getPPr().getOutlineLvl().getVal());

                return String.valueOf(
                        doc.getStyles().getStyle(para.getStyle()).getCTStyle().getPPr().getOutlineLvl().getVal());
            }
        } catch (Exception e) {

        }

        try {
            // 判断该段落的样式的基础样式是否设置了大纲级别
            if (doc.getStyles().getStyle(doc.getStyles().getStyle(para.getStyle()).getCTStyle().getBasedOn().getVal())
                    .getCTStyle().getPPr().getOutlineLvl() != null) {
                // System.out.println("getBasedOn");
                // System.out.println(para.getParagraphText());
                String styleName = doc.getStyles().getStyle(para.getStyle()).getCTStyle().getBasedOn().getVal();
                // System.out.println(doc.getStyles().getStyle(styleName).getCTStyle().getPPr().getOutlineLvl().getVal());

                return String
                        .valueOf(doc.getStyles().getStyle(styleName).getCTStyle().getPPr().getOutlineLvl().getVal());
            }
        } catch (Exception e) {

        }

        try {
            if (para.getStyleID() != null) {
                return para.getStyleID();
            }
        } catch (Exception e) {

        }

        return titleLvl;
    }

    /**
     * 获取标题编号
     *
     * @param titleLvl
     * @return
     */
    public static String getOrderCode(String titleLvl) {
        Map<String, Map<String, Object>> orderMap = new HashMap<String, Map<String, Object>>();
        String order = "";

        if (StringUtils.isEmpty(titleLvl) || "0".equals(titleLvl) || Integer.parseInt(titleLvl) == 8) {// 文档标题||正文
            return "";
        } else if (Integer.parseInt(titleLvl) > 0 && Integer.parseInt(titleLvl) < 8) {// 段落标题

            // 设置最高级别标题
            Map<String, Object> maxTitleMap = orderMap.get("maxTitleLvlMap");
            if (null == maxTitleMap) {// 没有，表示第一次进来
                // 最高级别标题赋值
                maxTitleMap = new HashMap<String, Object>();
                maxTitleMap.put("lvl", titleLvl);
                orderMap.put("maxTitleLvlMap", maxTitleMap);
            } else {
                String maxTitleLvl = maxTitleMap.get("lvl") + "";// 最上层标题级别(0,1,2,3)
                if (Integer.parseInt(titleLvl) < Integer.parseInt(maxTitleLvl)) {// 当前标题级别更高
                    maxTitleMap.put("lvl", titleLvl);// 设置最高级别标题
                    orderMap.put("maxTitleLvlMap", maxTitleMap);
                }
            }

            // 查父节点标题
            int parentTitleLvl = Integer.parseInt(titleLvl) - 1;// 父节点标题级别
            Map<String, Object> cMap = orderMap.get(titleLvl);// 当前节点信息
            Map<String, Object> pMap = orderMap.get(parentTitleLvl + "");// 父节点信息

            if (0 == parentTitleLvl) {// 父节点为文档标题，表明当前节点为1级标题
                int count = 0;
                // 最上层标题，没有父节点信息
                if (null == cMap) {// 没有当前节点信息
                    cMap = new HashMap<String, Object>();
                } else {
                    count = Integer.parseInt(String.valueOf(cMap.get("cCount")));// 当前序个数
                }
                count++;
                order = count + "";
                cMap.put("cOrder", order);// 当前序
                cMap.put("cCount", count);// 当前序个数
                orderMap.put(titleLvl, cMap);

            } else {// 父节点为非文档标题
                int count = 0;
                // 如果没有相邻的父节点信息，当前标题级别自动升级
                if (null == pMap) {
                    return getOrderCode(String.valueOf(parentTitleLvl));
                } else {
                    String pOrder = String.valueOf(pMap.get("cOrder"));// 父节点序
                    if (null == cMap) {// 没有当前节点信息
                        cMap = new HashMap<String, Object>();
                    } else {
                        count = Integer.parseInt(String.valueOf(cMap.get("cCount")));// 当前序个数
                    }
                    count++;
                    order = pOrder + "." + count;// 当前序编号
                    cMap.put("cOrder", order);// 当前序
                    cMap.put("cCount", count);// 当前序个数
                    orderMap.put(titleLvl, cMap);
                }
            }

            // 字节点标题计数清零
            int childTitleLvl = Integer.parseInt(titleLvl) + 1;// 子节点标题级别
            Map<String, Object> cdMap = orderMap.get(childTitleLvl + "");//
            if (null != cdMap) {
                cdMap.put("cCount", 0);// 子节点序个数
                orderMap.get(childTitleLvl + "").put("cCount", 0);
            }
        }
        return order;
    }

}
