package Bookmark;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageTree;
import org.apache.pdfbox.pdmodel.interactive.action.PDActionGoTo;
import org.apache.pdfbox.pdmodel.interactive.documentnavigation.destination.PDPageXYZDestination;
import org.apache.pdfbox.pdmodel.interactive.documentnavigation.outline.PDDocumentOutline;
import org.apache.pdfbox.pdmodel.interactive.documentnavigation.outline.PDOutlineItem;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Scanner;

public class AddBookmarks {

    public static void main(String[] args) throws IOException {

        String prefix = "18214";
        String searchKeyword = "ROUTINE CARD";



        ArrayList<String> filePathArray = new ArrayList<>();
        Scanner scanner = new Scanner(System.in);


        String excelPath = "C:\\Users\\mitha\\OneDrive\\Desktop\\acrobat-sign-main\\sdks\\PdfBox\\src\\main\\java\\Bookmark\\BookmarkDirectories.xlsx";
        XSSFWorkbook workbook = new XSSFWorkbook(excelPath);
        XSSFSheet sheet = workbook.getSheet("Sheet1");

        int rowCount = sheet.getPhysicalNumberOfRows();


        for (int i = 0; i < rowCount; i++) {

            String inputFilePath = sheet.getRow(i).getCell(0).getStringCellValue();

            filePathArray.add(inputFilePath);
        }


        for (String inputFilePath : filePathArray) {

            System.out.println(inputFilePath);
            System.out.println("Enter item start number");
            int startNumber = scanner.nextInt();
            scanner.nextLine();

            String outputFilePath = inputFilePath + "_Bookmarked.pdf";

            try {
                System.out.println("Loading the input PDF file...");
                PDDocument document = PDDocument.load(new File(inputFilePath));
                PDPageTree pages = document.getPages();
                PDDocumentOutline outline = new PDDocumentOutline();
                document.getDocumentCatalog().setDocumentOutline(outline);

                PDFTextStripper stripper = new PDFTextStripper();

                for (int pageIndex = 0; pageIndex < pages.getCount(); pageIndex++) {
                    System.out.println("Processing page " + (pageIndex + 1) + "...");
                    PDPage page = pages.get(pageIndex);
                    stripper.setStartPage(pageIndex + 1);
                    stripper.setEndPage(pageIndex + 1);
                    String pageText = stripper.getText(document);


                    if (pageText.toLowerCase().contains(searchKeyword.toLowerCase())) {
                        String bookmarkName = prefix + "-" + String.format("%04d", startNumber++);
                        System.out.println("Adding bookmark: " + bookmarkName);
                        PDOutlineItem bookmark = new PDOutlineItem();
                        bookmark.setTitle(bookmarkName);
                        PDPageXYZDestination dest = new PDPageXYZDestination();
                        dest.setPage(page);
                        dest.setZoom(1); // Adjust the zoom level as needed.
                        dest.setTop(750); // Adjust the vertical position as needed.
                        PDActionGoTo action = new PDActionGoTo();
                        action.setDestination(dest);
                        bookmark.setAction(action);
                        outline.addLast(bookmark);
                    }
                }

                System.out.println("Saving the output PDF file...");
                document.save(outputFilePath);
                document.close();
                System.out.println("Output PDF file saved successfully.");
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }
}