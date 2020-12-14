package com.excelTest.demo;

import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.servlet.ModelAndView;

import java.util.ArrayList;
import java.util.List;

@org.springframework.stereotype.Controller
public class Controller {

    @GetMapping(value = "/export")
    public ModelAndView getExcelFile(){
        List<Order> list = new ArrayList<Order>();
        list.add(new Order("abdulaziz", "alharthi", "26", "a@gmail.com"));
        list.add(new Order("abdulaziz", "alharthi", "26", "a@gmail.com"));
        list.add(new Order("abdulaziz", "alharthi", "26", "a@gmail.com"));
        list.add(new Order("abdulaziz", "alharthi", "26", "a@gmail.com"));
        list.add(new Order("abdulaziz", "alharthi", "26", "a@gmail.com"));
        return new ModelAndView(new OrderExcelSheet(), "list", list);
    }
}
