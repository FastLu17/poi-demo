package com.lxf.poi;

import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

@RunWith(SpringRunner.class)
@SpringBootTest
public class PoiApplicationTests {

    /**
     *  简单使用Java的正则表达式
     */
    @Test
    public void testMatcher() {
        Pattern pattern = Pattern.compile("\\$\\{(.+?)}", Pattern.CASE_INSENSITIVE);
        Matcher matcher = pattern.matcher("${name}---${age}");

        while (matcher.find()) {//正常都是使用while()进行迭代获取、
            System.out.println("groupCount = " + matcher.groupCount());
            System.out.println("group(1) = " + matcher.group(1));
            System.out.println("group(0) = " + matcher.group());
        }

        if (matcher.find()) {
            System.out.println("第一次匹配");
            System.out.println("groupCount = " + matcher.groupCount());
            System.out.println("group(1) = " + matcher.group(1));
            System.out.println("group(0) = " + matcher.group());
        }
        if (matcher.find()) {
            System.out.println("第二次匹配");
            System.out.println("groupCount = " + matcher.groupCount());
            System.out.println("group(1) = " + matcher.group(1));
            System.out.println("group(0) = " + matcher.group());
        }

    }

}
