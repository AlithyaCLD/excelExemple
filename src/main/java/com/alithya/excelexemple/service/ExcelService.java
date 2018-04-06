package com.alithya.excelexemple.service;

import java.io.File;
import java.io.IOException;
import java.util.Optional;
import java.util.stream.IntStream;

import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.tomcat.util.http.fileupload.ByteArrayOutputStream;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.ClassPathResource;
import org.springframework.stereotype.Service;

import com.alithya.excelexemple.config.ApplicationProperties;
import com.itextpdf.text.Document;
import com.itextpdf.text.Phrase;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;

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
    
    public byte[] calculateAndExportToPDF(String from, String to, String optionA, String optionB, String optionC) throws Exception {
        File excelFile = new ClassPathResource(properties.getExcelFileName()).getFile();
        log.debug("Chemin du fichier excel: {}", excelFile.getPath());

        Workbook workbook = new XSSFWorkbook(excelFile);

        Sheet sheet = workbook.getSheetAt(0);
        setCellValue(sheet, "A1", optionA);
        setCellValue(sheet, "B1", optionB);
        setCellValue(sheet, "C1", optionC);

        workbook.getCreationHelper().createFormulaEvaluator().evaluateAll();
        byte[] result = excelToPdf(sheet, from, to);

        workbook.close();

        return result;
    }

    private byte[] excelToPdf(Sheet sheet, String rangeStart, String rangeEnd) throws Exception {
        try {
            CellReference start = new CellReference(rangeStart);
            CellReference end = new CellReference(rangeEnd);
            int totalColumns = end.getCol() - start.getCol() + 1;

            // We will create output PDF document objects at this point
            Document pdfDocument = new Document();
            ByteArrayOutputStream outputStream = new ByteArrayOutputStream();

            PdfWriter.getInstance(pdfDocument, outputStream);
            pdfDocument.open();

            PdfPTable table = new PdfPTable(totalColumns);

            // Loop through rows and cells
            IntStream rows = IntStream.range(start.getRow(), end.getRow()+1);
            rows.forEach(rowIndex -> {
                Row row = sheet.getRow(rowIndex);

                if (row != null) {
                    IntStream.range(start.getCol(), end.getCol()+1).forEach(cellIndex -> {
                        Cell cell = row.getCell(cellIndex, MissingCellPolicy.CREATE_NULL_AS_BLANK);

                        switch (cell.getCellTypeEnum()) {
                        case BLANK:
                        case STRING:
                            PdfPCell pdfTextCell = new PdfPCell(new Phrase(cell.getStringCellValue()));
                            table.addCell(pdfTextCell);
                            break;
                        case FORMULA:
                        case NUMERIC:
                            PdfPCell pdfNumericCell = new PdfPCell(new Phrase(String.valueOf(cell.getNumericCellValue())));
                            table.addCell(pdfNumericCell);
                            break;
                        default:
                            log.warn("Celule avec un type pas traité: {} - {}", cell.getAddress(), cell.getCellTypeEnum());
                            break;
                        }
                    });
                }

                table.completeRow();
            });

            // Finally add the table to PDF document
            pdfDocument.add(table);
            pdfDocument.close();

            outputStream.flush();
            return outputStream.toByteArray();
        } catch (Exception e) {
            log.error("Unexpected error while converting to PDF", e);
            throw e;
        }
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
