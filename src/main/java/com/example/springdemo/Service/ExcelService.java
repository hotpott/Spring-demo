package com.example.springdemo.Service;

import com.example.springdemo.pojo.ExcelPojo;

import java.io.IOException;

/**
 * Excel相关服务
 *
 * @author gaoyd-a
 * @date 2023/10/24
 */
public interface ExcelService {

    /**
     *
     */
    ExcelPojo read() throws IOException;
}
