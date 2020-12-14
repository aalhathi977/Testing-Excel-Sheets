package com.excelTest.demo;


import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.web.servlet.view.document.AbstractXlsView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.util.Map;

public class OrderExcelSheet extends AbstractXlsView {

    @Override
    protected void buildExcelDocument(Map<String, Object> map, Workbook workbook, HttpServletRequest httpServletRequest, HttpServletResponse httpServletResponse) throws Exception {
        httpServletResponse.setContentType("application/ms-excel");
        httpServletResponse.setHeader("Content-Disposition", "attachment;filename=\"orders.xls\"");
        Iterable<Order> list = (Iterable<Order>) map.get("list");

        Sheet sheet = workbook.createSheet("Orders");
        sheet.setDefaultColumnWidth((short) 20);
        Row header = sheet.createRow(0);

        header.createCell(0).setCellValue("First Name");
        header.createCell(1).setCellValue("Last Name");
        header.createCell(2).setCellValue("Age");
        header.createCell(3).setCellValue("Email");

        int rowNum = 1;
        for(Order order: list){
            Row row = sheet.createRow(rowNum++);
            row.createCell(0).setCellValue(order.getFname());
            row.createCell(1).setCellValue(order.getLname());
            row.createCell(2).setCellValue(order.getAge());
            row.createCell(3).setCellValue(order.getEmail());
        }

    }
}
