package com.example.demo.service.implementation;

import com.example.demo.model.PdfData;
import com.example.demo.service.IExcelService;
import com.example.demo.service.IPdfGeneratorService;
import com.example.demo.service.IWordService;
import com.example.demo.util.CommonUtils;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;
import java.util.Map;

@Service
public class PdfGeneratorService implements IPdfGeneratorService {
    private static final String TEMPLATE_PATH = "D:/myFolder/template.docx"; // Template location
    private static final String OUTPUT_DIRECTORY = "D:/myFolder/result/";
    private static final String OUTPUT_EXTENSION_DOCX = ".docx";
    private static final String OUTPUT_EXTENSION_PDF = ".pdf";

    private final IWordService wordService;
    private final IExcelService excelService;

    public PdfGeneratorService(IExcelService excelService, IWordService wordService) {
        this.excelService = excelService;
        this.wordService = wordService;
    }
    public static void checkAndCreateDirectory(String outputDirectory) {
        Path path = Paths.get(outputDirectory);

        // Check if the directory exists
        if (Files.notExists(path)) {
            try {
                // If the directory doesn't exist, create it
                Files.createDirectories(path);
                System.out.println("Directory created: " + outputDirectory);
            } catch (IOException e) {
                System.err.println("Failed to create directory: " + e.getMessage());
            }
        } else {
            System.out.println("Directory already exists: " + outputDirectory);
        }
    }

    @Override
    public void generatePdfs(List<PdfData> pdfDataList)  {
        for (PdfData data : pdfDataList) {
            this.generateSinglePdf(data);
        }
    }

    @Override
    public void processExcelAndGeneratePdfs(MultipartFile excelFile) throws Exception {
        List<PdfData> excelData = excelService.readExcelFilePdfData(excelFile);
        generateMultiplePdfs(excelData); //Generate PDF file for each row
    }

    private void generateMultiplePdfs(List<PdfData> pdfDataList) {
        for (PdfData data : pdfDataList) {
            convertWordToPdf(data);
        }
    }

    private void generateSinglePdf(PdfData data) {
        convertWordToPdf(data);
    }

    private void convertWordToPdf(PdfData data) {
        String fileName = CommonUtils.generateFileName(data.getSt_full_name()) + OUTPUT_EXTENSION_PDF;
        checkAndCreateDirectory(OUTPUT_DIRECTORY);
        try {
            String outputPdfPath = OUTPUT_DIRECTORY + fileName;
            Map<String, String> record = CommonUtils.createRecordFromPdfData(data);
            wordService.convertWordToPDF(record, TEMPLATE_PATH, outputPdfPath);
            System.out.println("Generated PDF: " + outputPdfPath);
        } catch (Exception e) {
            System.err.println("Error generating PDF: " + fileName);
            e.printStackTrace();
        }
    }

    public void generateDocxs(List<PdfData> pdfDataList)  {
        for (PdfData data : pdfDataList) {
            this.generateSingleDoc(data);
        }
    }

    public void processExcelAndGenerateDocxs(MultipartFile excelFile) throws Exception {
        List<PdfData> excelData = excelService.readExcelFilePdfData(excelFile);
        generateMultipleDocxs(excelData); //Generate PDF file for each row
    }

    private void generateMultipleDocxs(List<PdfData> pdfDataList) {
        for (PdfData data : pdfDataList) {
            convertWordToDocx(data);
        }
    }

    private void generateSingleDoc(PdfData data) {
        convertWordToDocx(data);
    }
    private void convertWordToDocx(PdfData data) {
        String fileName = CommonUtils.generateFileName(data.getSt_full_name()) + OUTPUT_EXTENSION_DOCX;
        checkAndCreateDirectory(OUTPUT_DIRECTORY);
        try {
            String outputPdfPath = OUTPUT_DIRECTORY + fileName;
            Map<String, String> record = CommonUtils.createRecordFromPdfData(data);
            wordService.generateWordFile(record, TEMPLATE_PATH, outputPdfPath);
            System.out.println("Generated DOCX: " + outputPdfPath);
        } catch (Exception e) {
            System.err.println("Error generating DOCX: " + fileName);
            e.printStackTrace();
        }
    }

}

