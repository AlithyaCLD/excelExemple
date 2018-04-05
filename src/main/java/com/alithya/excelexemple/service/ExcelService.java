package com.alithya.excelexemple.service;

import java.io.File;
import java.io.IOException;
import java.util.Optional;

import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.ClassPathResource;
import org.springframework.stereotype.Service;

import com.alithya.excelexemple.config.ApplicationProperties;

import lombok.extern.slf4j.Slf4j;

@Slf4j
@Service
public class ExcelService {

    private static final String DEFAULT_CALCULATION_CELL = "D1";
    
	@Autowired
	private ApplicationProperties properties;
	
    public String calculate(String optionA, String optionB, String optionC) {
        return calculateUsing(DEFAULT_CALCULATION_CELL, optionA, optionB, optionC);
    }

    public String calculateUsing(String cell, String optionA, String optionB, String optionC) {
        Optional<String> origin = Optional.of(cell);
        String result = null;

        try {
            File excelFile = new ClassPathResource(properties.getExcelFileName()).getFile();
            log.debug("Chemin du fichier excel: {}", excelFile.getPath());

            Workbook workbook = new XSSFWorkbook(excelFile);

            Sheet sheet = workbook.getSheetAt(0);
            setCellValue(sheet, "A1", optionA);
            setCellValue(sheet, "B1", optionB);
            setCellValue(sheet, "C1", optionC);

            workbook.getCreationHelper().createFormulaEvaluator().evaluateAll();

            CellReference cellReference = getCellReference(origin.orElse(DEFAULT_CALCULATION_CELL), workbook);
            result = getCellValueAsString(cellReference, sheet);

            workbook.close();
        } catch (IOException | InvalidFormatException e) {
            log.error("Erreur pour ouvrir le fichier Excel.", e);
        }

        return result;
    }

    private String getCellValueAsString(CellReference reference, Sheet sheet) {
        Row row = sheet.getRow(reference.getRow());
        Cell cell = row.getCell(reference.getCol());

        return String.valueOf(cell.getNumericCellValue());
    }

    private CellReference getCellReference(String reference, Workbook workbook) {
        CellReference cellReference;

        int nameIndex = workbook.getNameIndex(reference);
        if (nameIndex >= 0) {
            cellReference = new CellReference(workbook.getNameAt(nameIndex).getRefersToFormula());
        } else {
            cellReference = new CellReference(reference);
        }

        return cellReference;
    }

	private void setCellValue(Sheet sheet, String reference, String value) {
		CellReference cellReference = new CellReference(reference);
		
		Row row = sheet.getRow(cellReference.getRow());
		Cell cell = row.getCell(cellReference.getCol());
		
		if (cell == null) {
			cell = row.createCell(cellReference.getCol());
		}

		cell.setCellValue(value);
	}

}
