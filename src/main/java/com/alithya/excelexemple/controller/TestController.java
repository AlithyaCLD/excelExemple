package com.alithya.excelexemple.controller;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseStatus;

import com.alithya.excelexemple.service.ExcelService;

@Controller
public class TestController {

    private ExcelService excelService;
    
    @Autowired
    public TestController(ExcelService excelService) {
        this.excelService = excelService;
    }

    @GetMapping(path = "/calculate")
    @ResponseStatus(HttpStatus.OK)
    public ResponseEntity<String> calculate(@RequestParam(name="optionA") String optionA, 
    		@RequestParam(name="optionB") String optionB, 
    		@RequestParam(name="optionC") String optionC) {
    	return new ResponseEntity<>(excelService.calculate(optionA, optionB, optionC), HttpStatus.OK);
    }
    
}
