package Bookmark;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.PDPageTree;
import org.apache.pdfbox.pdmodel.font.PDType1Font;
import org.apache.pdfbox.pdmodel.interactive.action.PDActionGoTo;
import org.apache.pdfbox.pdmodel.interactive.documentnavigation.destination.PDPageDestination;
import org.apache.pdfbox.pdmodel.interactive.documentnavigation.destination.PDPageXYZDestination;
import org.apache.pdfbox.pdmodel.interactive.documentnavigation.outline.PDOutlineItem;
import org.apache.pdfbox.pdmodel.interactive.documentnavigation.outline.PDDocumentOutline;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.awt.*;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class WriteExistingBookmark {

    public static void main(String[] args) throws IOException {
        ArrayList<String> filePathArray = new ArrayList<>();

        String excelPath = "C:\\Users\\mitha\\IdeaProjects\\PdfBox\\src\\main\\java\\Bookmark\\BookmarkDirectories.xlsx";
        XSSFWorkbook workbook = new XSSFWorkbook(excelPath);
        XSSFSheet sheet = workbook.getSheet("Sheet1");

        int rowCount = sheet.getPhysicalNumberOfRows();

        for (int j = 0; j < rowCount; j++) {
            String inputFilePath = sheet.getRow(j).getCell(0).getStringCellValue();
            filePathArray.add(inputFilePath);

            String outputFilePath = inputFilePath+"_BookmarksWritten.pdf";

            processPDF(inputFilePath, outputFilePath);

        }
    }

    private static void processPDF(String inputFilePath, String outputFilePath) throws IOException {
        System.out.println("Loading the input PDF file...");
        PDDocument document = PDDocument.load(new File(inputFilePath));
        PDPageTree pages = document.getPages();

        PDDocumentOutline outline = document.getDocumentCatalog().getDocumentOutline();
        if (outline != null) {
            List<PDOutlineItem> bookmarks = new ArrayList<>();
            gatherBookmarks(outline.children().iterator(), bookmarks);

            for (int i = 0; i < bookmarks.size(); i++) {
                PDOutlineItem bookmark = bookmarks.get(i);
                PDOutlineItem nextBookmark = (i < bookmarks.size() - 1) ? bookmarks.get(i + 1) : null;

                int startPage = getPageFromBookmark(bookmark, pages);
                int endPage = (nextBookmark != null) ? getPageFromBookmark(nextBookmark, pages) : pages.getCount();

                if (startPage != -1) {
                    for (int p = startPage; p < endPage; p++) {
                        writeBookmarkOnPage(bookmark.getTitle(), document, pages.get(p));
                    }
                }
            }
        } else {
            System.out.println("No bookmarks found in the PDF file.");
        }

        System.out.println("Saving the output PDF file...");
        document.save(outputFilePath);
        document.close();
        System.out.println("Output PDF file saved successfully.");
    }

    private static void gatherBookmarks(Iterator<PDOutlineItem> iterator, List<PDOutlineItem> bookmarks) {
        while (iterator.hasNext()) {
            PDOutlineItem bookmark = iterator.next();
            bookmarks.add(bookmark);
            if (bookmark.children().iterator().hasNext()) {
                gatherBookmarks(bookmark.children().iterator(), bookmarks);
            }
        }
    }

    private static int getPageFromBookmark(PDOutlineItem bookmark, PDPageTree pages) throws IOException {
        PDActionGoTo action = (PDActionGoTo) bookmark.getAction();
        if (action != null) {
            PDPageDestination dest = (PDPageDestination) action.getDestination();
            if (dest instanceof PDPageXYZDestination) {
                PDPage page = dest.getPage();
                if (page != null) {
                    return pages.indexOf(page);
                }
            }
        }
        return -1;
    }

    private static void writeBookmarkOnPage(String bookmarkTitle, PDDocument document, PDPage page) throws IOException {
        PDPageContentStream contentStream = new PDPageContentStream(document, page, PDPageContentStream.AppendMode.APPEND, true);
        contentStream.setFont(PDType1Font.HELVETICA_BOLD, 10);
        contentStream.setNonStrokingColor(Color.RED);
        contentStream.beginText();
        contentStream.newLineAtOffset(475, 828); //the current setting writes it to bottom right. tx:494 and ty:840 write on the top right of the page
        contentStream.showText(bookmarkTitle);
        contentStream.endText();
        contentStream.close();
    }
}
