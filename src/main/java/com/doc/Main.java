package com.doc;

import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.util.StringUtil;

import java.security.Guard;
import java.util.*;

public class Main {

    public static void main(String[] args) {
        // write your code here

        String path = "D://print//docMachine.docx";

        LiveDoc doc = new LiveDoc(path);
        Map<String, Object> map = new HashMap<>();
        map.put("title", "测试文书记录");
        map.put("same", 1);
        map.put("nosame", "无说明");
        map.put("parson", 6);
        map.put("prodate","2019-10-10");
        map.put("proname","东莞生产总企业");
        doc.setLabel(map);

        List<Map<String, Object>> table1 = new ArrayList<>();
        Map<String, Object> tableMap1 = new HashMap<>();
        tableMap1.put("name", "陈先生");
        tableMap1.put("date", "2020");
        tableMap1.put("code", "代码");
        table1.add(tableMap1);

        tableMap1.put("name", "何先生");
        tableMap1.put("date", "2019");
        tableMap1.put("code", "代码2");
        table1.add(tableMap1);

        doc.setTable(table1, "firstTable");

        doc.save("D://print//docx//" + UUID.randomUUID() + ".docx");

        System.out.print("doc machine, doc creat succes");
    }
}
