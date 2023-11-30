package com.example.springdemo.pojo;

import lombok.Data;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

/**
 * TODO
 *
 * @author gaoyd-a
 * @date 2023/10/24
 */
@Data
public class SheetPojo {

    /**
     * sheet页标
     */
    int sheetNum;

    /**
     * sheet名
     */
    String sheetName;

    /**
     * 标题行
     */
    List<String> title;

    /**
     * 数据
     */
    List<Map<String,String>> data;



    public List<String> getByColumn(String titleName) {
        if (!this.title.contains(titleName)) {
            return new ArrayList<>();
        }
        return data.stream().map(a -> a.get(titleName)).collect(Collectors.toList());
    }
}
