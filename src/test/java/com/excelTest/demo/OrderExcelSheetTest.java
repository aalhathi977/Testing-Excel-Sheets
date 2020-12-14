package com.excelTest.demo;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.autoconfigure.web.servlet.WebMvcTest;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.web.WebAppConfiguration;
import org.springframework.test.web.servlet.MockMvc;
import org.springframework.test.web.servlet.MvcResult;
import org.springframework.test.web.servlet.request.MockMvcRequestBuilders;
import org.springframework.test.web.servlet.result.MockMvcResultHandlers;
import org.springframework.test.web.servlet.result.MockMvcResultMatchers;
import org.springframework.test.web.servlet.setup.MockMvcBuilders;
import org.springframework.web.context.WebApplicationContext;

import java.io.ByteArrayInputStream;

import static org.junit.jupiter.api.Assertions.assertEquals;

@WebMvcTest(Controller.class)
class OrderExcelSheetTest {

    @Autowired
    private MockMvc mockMvc;


    @Test
    public void testGeneratedExcelSheet() throws Exception {
        MvcResult mvcResult = this.mockMvc.perform(
                MockMvcRequestBuilders.get("/export"))
                .andDo(MockMvcResultHandlers.print())
                .andExpect(MockMvcResultMatchers.status().isOk())
                .andReturn();
        byte[] workbookBytes = mvcResult.getResponse().getContentAsByteArray();

        Workbook actual = new HSSFWorkbook(new ByteArrayInputStream(workbookBytes));
        String actualValue = actual
                .getSheetAt(0)
                .getRow(0)
                .getCell(0)
                .getStringCellValue();
        assertEquals("First Name", actualValue);
    }

}