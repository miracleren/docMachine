package com.doc;



import java.io.*;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;


import org.apache.poi.xwpf.usermodel.*;

import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;


public class LiveDoc {

    /*
    read me
    run分割时{{}}有bug，{{}}标签块内容必从文本工具复制粘贴过来，或使用word替换工具把内容替换一次，才能被识别为一个run
    {{}}标签如果结尾处为换行符，也会出现无法识别，可在换行结尾添加白色符号
     */

    private static final String POI_MINIMUM_VERSION = "4.0.0";
    private XWPFDocument doc = new XWPFDocument();

    public LiveDoc(String filePath) {
        try {
            InputStream in = new FileInputStream(filePath);
            //FileInputStream in = new FileInputStream(filePath);
            doc = new XWPFDocument(in);
            in.close();
            System.out.println("create docx written successully");
        } catch (Exception e) {
            System.out.println(e.toString());
        }

    }


    /**
     * 标签数据替换
     *
     * @param values map值列表
     */
    public void setLabel(Map<String, Object> values) {
        List<XWPFParagraph> paras = doc.getParagraphs();
        List<XWPFRun> runs;
        Matcher matcher;
        //循环文档内容
        for (XWPFParagraph para : paras) {
            if (matcher(para.getParagraphText()).find()) {
                runs = para.getRuns();
                for (XWPFRun run : runs) {
                    String runText = run.toString();
                    matcher = matcher(runText);
                    System.out.println(runText);
                    if (matcher.find()) {
                        String par = matcher.group();
                        String[] parCut = par.split(":");
                        if (parCut.length == 1) {
                            //纯标签替换
                            String labVal = StringOf(values.get(par));
                            run.setText(StringOf(values.get(par)), 0);
                        } else if (parCut.length == 3) {
                            //类型标签
                            String typeName = parCut[0];
                            String typePar = parCut[1];
                            String typeVal = parCut[2];
                            if ("SC".equals(typeName)) {
                                //单选
                                if (String.valueOf(values.get(typePar)).equals(typeVal))
                                    run.setText("√", 0);
                                else
                                    run.setText("□", 0);
                            } else if ("MC".equals(typeName)) {//多选
                                int parval = Integer.parseInt(values.get(typePar).toString());
                                int val = Integer.parseInt(typeVal);
                                if ((parval & val) == val)
                                    run.setText("√", 0);
                                else
                                    run.setText("□", 0);
                            }
                        }

                    }
                }
            }
        }

        //循环文档表格
        List<XWPFTable> tableList = doc.getTables();
        for (XWPFTable table : tableList) {
            for (XWPFTableRow row : table.getRows()) {
                for (XWPFTableCell cell : row.getTableCells()) {
                    for (XWPFParagraph grap : cell.getParagraphs()) {
                        for (XWPFRun run : grap.getRuns()) {
                            String runText = run.toString();
                            String par = getFirstParName(runText);
                            if (!par.equals("") && !par.contains(":")) {
                                runText = runText.replace("{{" + par + "}}", StringOf(values.get(par)));
                                run.setText(runText, 0);
                            }
                        }
                    }
                }
            }
        }
    }

    /**
     * 循环填充表格内容
     */
    public void setTable(List<Map<String, Object>> list, String tableName) {
        List<XWPFTable> tableList = doc.getTables();
        for (XWPFTable table : tableList) {
            //寻找表格配置信息
            XWPFTableRow firstRow = table.getRow(0);
            String tableConfig = getFirstParName(firstRow.getCell(0).getText());
            if (!tableConfig.equals("")) {
                //第一行为表格配置信息
                String[] cons = tableConfig.split(":");
                if (!cons[0].equals("TABLE") && !cons[1].equals(tableName))
                    break;
            } else
                break;

            //查找循环列信息
            int i = 0, tempIndex = -1;
            XWPFTableRow tempRow = null;
            List<XWPFTableRow> trows = table.getRows();
            for (XWPFTableRow trow : trows) {
                if (tempRow != null)
                    break;
                for (XWPFTableCell tcell : trow.getTableCells()) {
                    if (getFirstParName(tcell.getText()).contains("COL")) {
                        tempRow = trow;
                        tempIndex = i;
                        break;
                    }
                }
                i++;
            }
            if (tempRow == null)
                return;

            //克隆行，并赋值
            for (Map<String, Object> rowData : list) {
                XWPFTableRow newRow = cloneRow(table, tempRow, i);
                for (XWPFTableCell rowCell : newRow.getTableCells()) {
                    List<XWPFParagraph> paragraphs = rowCell.getParagraphs();
                    for (XWPFParagraph xwpfParagraph : paragraphs) {
                        List<XWPFRun> runs = xwpfParagraph.getRuns();
                        for (XWPFRun run : runs) {
                            String runText = run.toString();
                            String parText = getFirstParName(runText);
                            if (!parText.equals("")) {
                                String[] pars = parText.split(":");
                                if (pars[0].equals("COL")) {
                                    runText = runText.replace("{{" + parText + "}}", StringOf(rowData.get(pars[1])));
                                    run.setText(runText, 0);
                                }
                            }
                        }
                    }


                }
                i++;
            }

            //删除循环信息列
            table.removeRow(tempIndex);
            //删除配置列
            table.removeRow(0);

        }
    }


    public XWPFTableRow cloneRow(XWPFTable table, XWPFTableRow sourceRow, int rowIndex) {
        //在表格指定位置新增一行
        XWPFTableRow targetRow = table.insertNewTableRow(rowIndex);

        //复制行属性
        if (sourceRow.getCtRow().getTrPr() != null)
            targetRow.getCtRow().setTrPr(sourceRow.getCtRow().getTrPr());

        List<XWPFTableCell> cellList = sourceRow.getTableCells();
        if (null != cellList) {
            //复制列及其属性和内容
            XWPFTableCell targetCell ;
            for (XWPFTableCell sourceCell : cellList) {
                targetCell = targetRow.addNewTableCell();
                //列属性
                targetCell.getCTTc().setTcPr(sourceCell.getCTTc().getTcPr());
                //段落属性
                if (sourceCell.getParagraphs() != null && sourceCell.getParagraphs().size() > 0) {
                    targetCell.getParagraphs().get(0).getCTP().setPPr(sourceCell.getParagraphs().get(0).getCTP().getPPr());
                    if (sourceCell.getParagraphs().get(0).getRuns() != null && sourceCell.getParagraphs().get(0).getRuns().size() > 0) {
                        XWPFRun cellR = targetCell.getParagraphs().get(0).createRun();
                        cellR.setText(sourceCell.getText());
                        cellR.setBold(sourceCell.getParagraphs().get(0).getRuns().get(0).isBold());
                    } else {
                        targetCell.setText(sourceCell.getText());
                    }
                } else {
                    targetCell.setText(sourceCell.getText());
                }
            }
        }
        return targetRow;
    }


    /**
     * {{par}} 参数查找正则
     *
     * @param str 查找串
     * @return  返结果
     */
    private Matcher matcher(String str) {
        Pattern pattern = Pattern.compile("(?<=\\{\\{)(.+?)(?=\\}\\})", Pattern.CASE_INSENSITIVE);
        Matcher matcher = pattern.matcher(str);
        return matcher;
    }

    /**
     * 获取数据里第一个标签名称
     *
     * @param str
     * @return
     */
    private String getFirstParName(String str) {
        Pattern pattern = Pattern.compile("(?<=\\{\\{)(.+?)(?=\\}\\})", Pattern.CASE_INSENSITIVE);
        Matcher matcher = pattern.matcher(str);
        if (matcher.find())
            return matcher.group();
        else
            return "";
    }

    private boolean isEmpty(String str) {
        return str == null || str.trim().equals("");
    }

    /**
     * 空字符转占位空格
     */
    private String StringOf(Object val) {
        return val == null ? "        " : val.toString();
    }

    /**
     * 保存文件到路径
     *
     * @param outPath 路径
     * @return
     */
    public boolean save(String outPath) {
        OutputStream outputStream;
        try {
            outputStream = new FileOutputStream(outPath);
            doc.validateProtectionPassword("test");
            doc.write(outputStream);
            doc.close();
            outputStream.close();
            return true;
        } catch (Exception e) {
            return false;
        }

    }

}
