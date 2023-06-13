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
import java.util.Collections;
import java.util.HashMap;
import java.util.Scanner;

public class WPBookmarksDuplicate {
    public static void main(String[] args) throws IOException {
        ArrayList<String> uploadList = new ArrayList<>();
        ArrayList<Integer> itemNumbers = new ArrayList<>(); // create itemNumbers ArrayList

        String prefix = "H901 WO 18483";

        String uploadListPath = "C:\\Users\\mitha\\OneDrive\\Desktop\\E1-23-05-00013_OE-ICA_6Y CHECK\\30_Wings\\OE-ICA TALLY2.xlsx";
        XSSFWorkbook uploadListWorkbook = new XSSFWorkbook(uploadListPath);
        XSSFSheet uploadSheet = uploadListWorkbook.getSheet("Sheet1");
        int rowCountUploadList = uploadSheet.getPhysicalNumberOfRows();
        String workOrderNumber = "";

        for (int i = 1; i < rowCountUploadList; i++) {
            workOrderNumber = uploadSheet.getRow(i).getCell(1).getStringCellValue();
            uploadList.add(workOrderNumber);


            int itemNumber = (int) uploadSheet.getRow(i).getCell(0).getNumericCellValue(); // read item number from column A
            itemNumbers.add(itemNumber); // add item number to itemNumbers ArrayList
        }

        ArrayList<String> filePathArray = new ArrayList<>();

        String excelPath = "C:\\Users\\mitha\\IdeaProjects\\PdfBox\\src\\main\\java\\Bookmark\\BookmarkDirectories.xlsx";
        XSSFWorkbook workbook = new XSSFWorkbook(excelPath);
        XSSFSheet sheet = workbook.getSheet("Sheet1");

        int rowCount = sheet.getPhysicalNumberOfRows();

        for (int j = 0; j < rowCount; j++) {
            String inputFilePath = sheet.getRow(j).getCell(0).getStringCellValue();
            filePathArray.add(inputFilePath);


        }


        if (!filePathArray.isEmpty()) {
            String inputFilePath = filePathArray.get(0);
            System.out.println(inputFilePath);

            String outputFilePath = inputFilePath + "_Bookmarked.pdf";

            try {
                processPDF(inputFilePath, outputFilePath, uploadList, itemNumbers, prefix);
            } catch (IOException e) {
                e.printStackTrace();
            }
        } else {
            System.out.println("No file paths found in the Excel file.");
        }
    }


    private static void processPDF(String inputFilePath, String outputFilePath, ArrayList<String> uploadList, ArrayList<Integer> itemNumbers, String prefix) throws IOException {
        System.out.println("Loading the input PDF file...");
        PDDocument document = PDDocument.load(new File(inputFilePath));
        PDPageTree pages = document.getPages();
        PDDocumentOutline outline = new PDDocumentOutline();
        document.getDocumentCatalog().setDocumentOutline(outline);

        PDFTextStripper stripper = new PDFTextStripper();
        System.out.println("Enter starting page");
        Scanner scanner = new Scanner(System.in);
        int currentPageIndex = scanner.nextInt() - 1;
        scanner.nextLine();

        // This map will keep track of how many bookmarks have been made for each keyword
        HashMap<String, Integer> keywordBookmarks = new HashMap<>();

        for (int i = 0; i < Math.min(uploadList.size(), itemNumbers.size()); i++) {
            String keyword = uploadList.get(i);
            int itemNumber = itemNumbers.get(i);

            int keywordBookmarksCount = keywordBookmarks.getOrDefault(keyword, 0);

            // Only process this keyword if we have fewer bookmarks than its count in uploadList
            if (Collections.frequency(uploadList, keyword) > keywordBookmarksCount) {

                for (int pageIndex = currentPageIndex; pageIndex < pages.getCount(); pageIndex++) {
                    System.out.println("Processing page " + (pageIndex + 1) + "...");
                    PDPage page = pages.get(pageIndex);
                    stripper.setStartPage(pageIndex + 1);
                    stripper.setEndPage(pageIndex + 1);
                    String pageText = stripper.getText(document);

                    if (pageText.toLowerCase().contains(keyword.toLowerCase())) {
                        String bookmarkName = prefix + "-" + String.format("%04d", itemNumber);
                        System.out.println("Adding bookmark: " + bookmarkName);
                        PDOutlineItem bookmark = new PDOutlineItem();
                        bookmark.setTitle(bookmarkName);
                        PDPageXYZDestination dest = new PDPageXYZDestination();
                        dest.setPage(page);
                        dest.setZoom(0.673F); // Adjust the zoom level as needed.
                        dest.setTop(1000); // Adjust the vertical position as needed.
                        PDActionGoTo action = new PDActionGoTo();
                        action.setDestination(dest);
                        bookmark.setAction(action);
                        outline.addLast(bookmark);

                        keywordBookmarks.put(keyword, keywordBookmarksCount + 1); // Increase the count of bookmarks for this keyword

                        currentPageIndex = pageIndex + 1;
                        break;
                    }
                }
            }
        }

        System.out.println("Saving the output PDF file...");
        document.save(outputFilePath);
        document.close();
        System.out.println("Output PDF file saved successfully.");
    }
}