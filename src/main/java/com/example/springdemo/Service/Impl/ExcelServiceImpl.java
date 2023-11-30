package com.example.springdemo.Service.Impl;

import com.example.springdemo.pojo.ExcelPojo;
import com.example.springdemo.pojo.SheetPojo;
import com.example.springdemo.Service.ExcelService;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.core.io.ClassPathResource;
import org.springframework.core.io.Resource;

import java.io.*;
import java.net.URLDecoder;
import java.util.*;

/**
 * EXCEL服务实现
 *
 * @author gaoyd-a
 * @date 2023/10/24
 */
@Slf4j
public class ExcelServiceImpl implements ExcelService {

    private final static String EXCEL_TEMPLATE_NAME = "GEPS_PMLead.xlsx";

    private final static Map<Integer, String> TITLE_MAP = new HashMap<>();

    @Override
    public ExcelPojo read() throws IOException {
        Resource resource = open();
        ExcelPojo excelEntity = new ExcelPojo();
        Workbook wb=null;
        if(resource==null) {
            return excelEntity;
        }
        try (InputStream inputStream = resource.getInputStream()) {
            if (EXCEL_TEMPLATE_NAME.endsWith(".xlsx")) {
                wb = new XSSFWorkbook(inputStream);
            } else if (EXCEL_TEMPLATE_NAME.endsWith(".xls")) {
                wb = new HSSFWorkbook(inputStream);
            } else {
                log.info("文件名错误");
            }
            int sheetNum = wb.getNumberOfSheets();
            for (int i=0; i<sheetNum;i++) {
                SheetPojo sheetEntity = new SheetPojo();
                Sheet sheet = wb.getSheetAt(i);
                sheetEntity.setSheetNum(i);
                sheetEntity.setSheetName(sheet.getSheetName());
                sheetEntity.setTitle(getTitleRow(sheet));
                sheetEntity.setData(readSheet(sheet));

                excelEntity.getNameSheetEntityMap().put(sheetEntity.getSheetName(), sheetEntity);
                excelEntity.getNameSheetEntityMap().put(sheetEntity.getSheetName(), sheetEntity);
            }
        } catch (IOException exception) {
            throw new IOException("文件读入失败");
        } finally {
            if (wb!=null) {
                wb.close();
            }
        }
        return excelEntity;
    }

    private Resource open() {
        try {
            Resource resource = new ClassPathResource(URLDecoder.decode(EXCEL_TEMPLATE_NAME, "utf-8"));
            return resource;
        } catch (IOException e) {
            e.printStackTrace();
        }
        return null;
    }
    private List<String> getTitleRow(Sheet sheet) {
        List<String> titleNameList = new ArrayList<>();
        Iterator iterator = sheet.rowIterator();
        if (iterator.hasNext()) {
            Row row = (Row) iterator.next();
            int colNum = row.getLastCellNum();
            for (int i = 0; i < colNum; i++) {
                Cell cell = row.getCell(i);
                if (cell != null) {
                    cell.setCellType(CellType.STRING);
                    String cellValue =  cell.getStringCellValue();
                    TITLE_MAP.put(i, cellValue);
                    titleNameList.add(cellValue);
                }
            }
        }
        return titleNameList;
    }

    private List<Map<String, String>> readSheet(Sheet sheet) {
        List<Map<String, String>> sheetDataList = new ArrayList<>();
        Iterator iterator = sheet.rowIterator();
        if (iterator.hasNext()) {
            // 跳过标题行
            iterator.next();
        }
        while (iterator.hasNext()) {
            Map<String, String> sheetData =  new HashMap<>(10);
            Row row = (Row) iterator.next();
            int colNum = row.getLastCellNum();
            for (int i = 0; i < colNum; i++) {
                Cell cell = row.getCell(i);
                if (cell != null) {
                    cell.setCellType(CellType.STRING);
                    sheetData.put(TITLE_MAP.get(i), cell.getStringCellValue());
                }
            }
            sheetDataList.add(sheetData);
        }
        return sheetDataList;
    }

    private void writeSheet(SheetPojo sheetEntity) {
        File file;
        FileOutputStream fileOutputStream;
        Workbook wb;
        try {
            Resource resource = new ClassPathResource(URLDecoder.decode(EXCEL_TEMPLATE_NAME, "utf-8"));
            if (!resource.isFile()) {
                throw new IOException("文件打开失败");

            }
            file = resource.getFile();
            fileOutputStream = new FileOutputStream(file);
            wb = new XSSFWorkbook(file);
            List<String> sheetNameList = getSheetNameList(wb);
            if (sheetNameList.contains(sheetEntity.getSheetName())) {

            }
        } catch (IOException exception) {
            log.info("写入失败");
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        }

    }

    private List<String> getSheetNameList(Workbook wb) {
        List<String> sheetNameList = new ArrayList<>();
        int numberOfSheet = wb.getNumberOfSheets();
        for (int i = 0; i < numberOfSheet; i++) {
            sheetNameList.add(wb.getSheetName(i));
        }
        return sheetNameList;
    }

    public static void main(String[] args){

        try {
            ExcelPojo excelEntity = new ExcelServiceImpl().read();
            SheetPojo gepsSheet = excelEntity.getSheet("GEPS");
            SheetPojo pmleadSheet = excelEntity.getSheet("PMLead");
            SheetPojo resSheet = excelEntity.getSheet("res");
            List<String> gepsComment = gepsSheet.getByColumn("注释");
            List<String> pmleadComment = pmleadSheet.getByColumn("注释");
            List<String> resComment = resSheet.getByColumn("table");
            List<String> resTableName = new ArrayList<>();
            List<String> resTableName2 = new ArrayList<>();
            for (String comment:gepsComment) {
                for (String res : resComment) {
                    if (res.contains(comment) && !resTableName.contains(res)) {
                        resTableName.add(res);
                        resTableName2.add(res + "_" + comment);
                    }
                }
            }
            log.info(resTableName.toString());

            Resource resource = new ClassPathResource(URLDecoder.decode(EXCEL_TEMPLATE_NAME, "utf-8"));
            Workbook wb = new XSSFWorkbook(resource.getInputStream());
            Sheet sheet = wb.getSheet("res");
            Iterator iterator = sheet.rowIterator();
            while (iterator.hasNext()) {
                Row row = (Row)iterator.next();
                Cell cell = row.getCell(2);
                cell.setCellType(CellType.STRING);
                if(resTableName.contains(cell.getStringCellValue())) {
                    row.createCell(5);
                    row.getCell(5).setCellValue("GEPS");
                }
            }
        } catch (IOException exception) {
            log.info("资源加载失败");
        }
    }


}
