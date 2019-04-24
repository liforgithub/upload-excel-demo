package com.lxy.service.impl;

import com.lxy.service.ResolveExcelService;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

@Service
public class ResolveExcelServiceImpl implements ResolveExcelService {

    private static final String SUFFIX_2003 = ".xls";
    private static final String SUFFIX_2007 = ".xlsx";

    @Override
    public void resolveExcel(MultipartFile file) throws Exception {

        if (file == null) {
            throw new Exception("对象不能为空");
        }
        //获取名字
        String originalFilename = file.getOriginalFilename();
        if (null == originalFilename) {
            throw new Exception("文件名为空");
        }
        Workbook workbook = null;
        try {
            if (originalFilename.endsWith(SUFFIX_2003)) {
                workbook = new HSSFWorkbook(file.getInputStream());
            } else if (originalFilename.endsWith(SUFFIX_2007)) {
                workbook = new XSSFWorkbook(file.getInputStream());
            }
        } catch (Exception e) {
            e.printStackTrace();
            throw new Exception("格式错误");
        }
        if (workbook == null) {
            throw new Exception("格式错误");
        } else {
            //获取所有的工作表的的数量
            int numOfSheet = workbook.getNumberOfSheets();
            //遍历这个这些表
            for (int i = 0; i < numOfSheet; i++) {
                //获取一个sheet也就是一个工作簿
                Sheet sheet = workbook.getSheetAt(i);
                int lastRowNum = sheet.getLastRowNum();
                //从第一行开始第一行一般是标题
                for (int j = 1; j <= lastRowNum; j++) {
                    Row row = sheet.getRow(j);

                    row.getCell(0).setCellType(Cell.CELL_TYPE_STRING);
                    String longName = row.getCell(0).getStringCellValue();
                    System.out.println(longName);
                }
            }
        }
    }
}
