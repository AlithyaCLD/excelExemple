package com.alithya.excelexemple.controller;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.core.io.Resource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
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
    public ResponseEntity<String> calculate(@RequestParam(name = "optionA") String optionA,
            @RequestParam(name = "optionB") String optionB,
            @RequestParam(name = "optionC") String optionC) {
        return new ResponseEntity<>(excelService.calculate(optionA, optionB, optionC), HttpStatus.OK);
    }

    @GetMapping(path = "/calculateUsing/{cell}")
    public ResponseEntity<String> calculateUsing(@PathVariable String cell,
            @RequestParam(name = "optionA") String optionA,
            @RequestParam(name = "optionB") String optionB,
            @RequestParam(name = "optionC") String optionC) {
        return new ResponseEntity<>(excelService.calculateUsing(cell, optionA, optionB, optionC), HttpStatus.OK);
    }

    @GetMapping(path = "/calculateAsPDF")
    public ResponseEntity<Resource> calculateAndExportToPDF(@RequestParam(name = "from") String from,
            @RequestParam(name = "to") String to,
            @RequestParam(name = "optionA") String optionA,
            @RequestParam(name = "optionB") String optionB,
            @RequestParam(name = "optionC") String optionC) throws Exception {
        byte[] result = excelService.calculateAndExportToPDF(from, to, optionA, optionB, optionC);

        HttpHeaders headers = new HttpHeaders();
        headers.add(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=excel.pdf");

        return ResponseEntity.ok()
                .headers(headers)
                .contentLength(result.length)
                .contentType(MediaType.APPLICATION_PDF)
                .body(new ByteArrayResource(result));
    }

}
