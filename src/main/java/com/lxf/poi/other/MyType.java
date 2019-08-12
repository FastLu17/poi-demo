package com.lxf.poi.other;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;

/**
 * 构造指定的类型、进行instanceof类型的判断
 *
 * @author 小66
 * @Description
 * @create 2019-08-12 10:46
 **/
public class MyType extends ArrayList<Map<String, Object>> {

    /**
     * 可以构造指定的类型、来进行 instanceof 类型的判断
     * <p>
     * public class MyType extends ArrayList<Map<String, Object>> {}
     *
     * @param params
     */
    private static void trimParamsByMyType(Map<String, Object> params) {
        for (Map.Entry<String, Object> entry : params.entrySet()) {
            if (entry.getValue() instanceof String) {
                params.put(entry.getKey(), entry.getValue().toString().trim());
            }
            if (entry.getValue() instanceof MyType) {
                MyType type = (MyType) entry.getValue();
                for (Map<String, Object> map : type) {
                    trimParamsByMyType(map);
                }
            }
        }
    }

    public static void main(String[] args) {
        Map<String, Object> params = new HashMap<>();
        params.put("a1", "aaa  aaa  ");
        params.put("a2", "  a222aa  aaa  ");
        params.put("a3", "   a333aa  aaa  ");
        Map<String, Object> map = new HashMap<>();
        map.put("b1", "bb bb  ");
        map.put("b2", "  b222b bb  ");
        map.put("b3", "    b3333b bb  ");
        //List<Map<String, Object>> mapList = new ArrayList<>();  不可用new出来的ArrayList、
        MyType myType = new MyType();
        myType.add(map);
        params.put("list", myType);
        System.out.println("params = " + params);
        trimParamsByMyType(params);
        System.out.println("params After = " + params);
    }
}
