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

public class WPBookmarks {
    public static void main(String[] args) throws IOException {
        ArrayList<String> uploadList = new ArrayList<>();
        ArrayList<String> itemNumbers = new ArrayList<>(); // create itemNumbers ArrayList

        String prefix = "";

        String uploadListPath = "C:\\Users\\martin.asenov\\Desktop\\A7-MHH Daily Reports\\Tally_bookmark.xlsx";
        XSSFWorkbook uploadListWorkbook = new XSSFWorkbook(uploadListPath);
        XSSFSheet uploadSheet = uploadListWorkbook.getSheet("Sheet1");
        int rowCountUploadList = uploadSheet.getPhysicalNumberOfRows();
        String workOrderNumber = "";

        for (int i = 1194  ; i <= 1222; i++) {  //rowCountUploadList
            workOrderNumber = uploadSheet.getRow(i).getCell(0).getStringCellValue();
            uploadList.add(workOrderNumber);


            String itemNumber = uploadSheet.getRow(i).getCell(1).getStringCellValue(); // read item number from column A
            itemNumbers.add(itemNumber); // add item number to itemNumbers ArrayList
        }

        ArrayList<String> filePathArray = new ArrayList<>();

        String excelPath = "C:\\Users\\martin.asenov\\IdeaProjects\\PdfBox\\src\\main\\java\\Bookmark\\BookmarkDirectories.xlsx";
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


    private static void processPDF(String inputFilePath, String outputFilePath, ArrayList<String> uploadList, ArrayList<String> itemNumbers, String prefix) throws IOException {
            System.out.println("Loading the input PDF file...");
            PDDocument document = PDDocument.load(new File(inputFilePath));
            PDPageTree pages = document.getPages();
            PDDocumentOutline outline = new PDDocumentOutline();
            document.getDocumentCatalog().setDocumentOutline(outline);

            PDFTextStripper stripper = new PDFTextStripper();


          System.out.println("Enter starting page");
          Scanner scanner=new Scanner(System.in);
          int currentPageIndex = scanner.nextInt()-1;
          scanner.nextLine();



        for (int i = 0; i < Math.min(uploadList.size(), itemNumbers.size()); i++) {

            String keyword = uploadList.get(i);
            String itemNumber = itemNumbers.get(i);



            for (int pageIndex = currentPageIndex; pageIndex < pages.getCount(); pageIndex++) {
                System.out.println("Processing page " + (pageIndex + 1) + "...");
                PDPage page = pages.get(pageIndex);
                stripper.setStartPage(pageIndex + 1);
                stripper.setEndPage(pageIndex + 1);
                String pageText = stripper.getText(document);

                if (pageText.toLowerCase().contains(keyword.toLowerCase())) {
                    String bookmarkName = prefix + "" + itemNumber;
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

                        currentPageIndex = pageIndex + 1;
                        break;
                    }
                }
            }

            System.out.println("Saving the output PDF file...");
            document.save(outputFilePath);
            document.close();
            System.out.println("Output PDF file saved successfully.");
        }
    }
