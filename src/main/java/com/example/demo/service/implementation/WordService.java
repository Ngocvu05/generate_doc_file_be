package com.example.demo.service.implementation;

import com.aspose.words.PdfCompliance;
import com.aspose.words.PdfSaveOptions;
import com.example.demo.service.IWordService;
import com.example.demo.util.CommonUtils;
import org.apache.poi.xwpf.usermodel.*;
import org.springframework.stereotype.Service;

import java.io.*;
import java.util.Map;

@Service
public class WordService implements IWordService {
    @Override
    public void convertWordToPDF(Map<String, String> data, String templatePath, String outputPdfPath) {
        try {
            XWPFDocument doc = replaceWordParam(data, templatePath);

            //Setting temporary docx file
            ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
            doc.write(byteArrayOutputStream);
            byte[] docxBytes = byteArrayOutputStream.toByteArray();

            //Convert docx to PDF
            com.aspose.words.Document asposeDoc = new com.aspose.words.Document(new java.io.ByteArrayInputStream(docxBytes));
            PdfSaveOptions saveOptions = new PdfSaveOptions(); { saveOptions.setCompliance(PdfCompliance.PDF_17); }
            asposeDoc.save(outputPdfPath);

            System.out.println("PDF created successfully at: " + outputPdfPath);
        } catch (Exception e) {
            e.printStackTrace();
            System.err.println("Error during DOCX to PDF conversion: " + e.getMessage());
        }
    }


    private  static XWPFDocument replaceWordParam(Map<String, String> data, String templatePath) throws IOException {
        // Load the Word template
        try (FileInputStream fis = new FileInputStream(templatePath)) {
            XWPFDocument doc = new XWPFDocument(fis);

            // Replace placeholders in paragraphs
            CommonUtils.replacePlaceholder(data, doc.getParagraphs());

            // Replace placeholders in tables (if the Word template contains tables)
            for (XWPFTable table : doc.getTables()) {
                for (XWPFTableRow row : table.getRows()) {
                    for (XWPFTableCell cell : row.getTableCells()) {
                        CommonUtils.replacePlaceholder(data, cell.getParagraphs());
                    }
                }
            }
            fis.close();
            return doc;
        }
    }

    public void generateWordFile(Map<String, String> data, String templatePath, String outputPath) throws Exception {
        // Load the Word template
        XWPFDocument doc = new XWPFDocument(new File(templatePath).toURI().toURL().openStream());

        doc = replaceWordParam(data, templatePath);

        // Save the generated Word file
        try (FileOutputStream out = new FileOutputStream(new File(outputPath))) {
            doc.write(out);
        }

        doc.close();
    }
}