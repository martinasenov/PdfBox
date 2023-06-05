package Bookmark;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

public class PageNumbers {

    public static void main(String[] args) throws IOException {

        ArrayList<String> uploadList = new ArrayList<>();

        String uploadListPath = "C:\\Users\\martin.asenov\\Desktop\\CAA BG\\Book1.xlsx";
        FileInputStream fis = new FileInputStream(uploadListPath);
        XSSFWorkbook uploadListWorkbook = new XSSFWorkbook(fis);
        XSSFSheet uploadSheet = uploadListWorkbook.getSheet("Sheet1");
        int rowCountUploadList = uploadSheet.getPhysicalNumberOfRows();
        String taskNumber = "";

        for (int i = 1; i < rowCountUploadList; i++) {
            taskNumber = uploadSheet.getRow(i).getCell(1).getStringCellValue();
            uploadList.add(taskNumber);
        }

        for (int i = 0; i < uploadList.size(); i++) {
            String task = uploadList.get(i);
            String filePath = "C:\\Users\\martin.asenov\\Desktop\\New folder\\WO8294_MPD WORKPACK_TC.pdf";

            try {
                int pageNumber = searchInPDF(filePath, task);
                System.out.println(pageNumber);

                Row row = uploadSheet.getRow(i + 1);
                if(row == null) {
                    row = uploadSheet.createRow(i + 1);
                }
                Cell cell = row.createCell(2);  // 2 represents column C
                cell.setCellValue(pageNumber);

            } catch (IOException e) {
                e.printStackTrace();
            }
        }

        // write changes to the file
        FileOutputStream fos = new FileOutputStream(uploadListPath);
        uploadListWorkbook.write(fos);
        fos.close();

        uploadListWorkbook.close();
        fis.close();
    }

    public static int searchInPDF(String fileName, String searchWord) throws IOException {
        PDDocument doc = PDDocument.load(new File(fileName));
        PDFTextStripper textStripper = new PDFTextStripper();

        int numberOfPages = doc.getNumberOfPages();

        for (int page = 1; page <= numberOfPages; page++) {
            textStripper.setStartPage(page);
            textStripper.setEndPage(page);
            String text = textStripper.getText(doc);
            if (text.contains(searchWord)) {
                doc.close();
                return page;
            }
        }
        doc.close();
        return -1;  // return -1 if the search word is not found in the PDF
    }
}