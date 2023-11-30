package com.example.springdemo.pojo;

import lombok.Data;

import java.util.HashMap;
import java.util.Map;

/**
 * TODO
 *
 * @author gaoyd-a
 * @date 2023/10/24
 */
@Data
public class ExcelPojo {

    Map<String, SheetPojo> nameSheetEntityMap = new HashMap<>();

    Map<Integer, SheetPojo> orderSheetEntityMap = new HashMap<>();

    public SheetPojo getSheet(String name) {
        return nameSheetEntityMap.get(name);
    }

    public SheetPojo getSheet(int order) {
        return orderSheetEntityMap.get(order);
    }
}
