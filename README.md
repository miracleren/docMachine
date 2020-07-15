# docMachine
基于POI实现模板生成word工具类


# 工具类使用方法
        //模板路径
        String path = "D://print//docMachine.docx";
        LiveDoc doc = new LiveDoc(path);
        
        //添加替换标签值，括单选、多选框（多选值采用2的平方值相加计算所得）
        Map<String, Object> map = new HashMap<>();
        map.put("title", "测试文书记录");
        map.put("same", 1);
        map.put("nosame", "无说明");
        map.put("parson", 6);
        map.put("prodate","2019-10-10");
        map.put("proname","东莞生产总企业");
        doc.setLabel(map);

        //动态生成表格，必需传递表格名称
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

