package Bookmark;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageTree;
import org.apache.pdfbox.pdmodel.interactive.action.PDActionGoTo;
import org.apache.pdfbox.pdmodel.interactive.documentnavigation.destination.PDDestination;
import org.apache.pdfbox.pdmodel.interactive.documentnavigation.destination.PDNamedDestination;
import org.apache.pdfbox.pdmodel.interactive.documentnavigation.destination.PDPageDestination;
import org.apache.pdfbox.pdmodel.interactive.documentnavigation.outline.PDDocumentOutline;
import org.apache.pdfbox.pdmodel.interactive.documentnavigation.outline.PDOutlineItem;
import org.apache.pdfbox.text.PDFTextStripper;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class RemoveDuplicates {
    public static void main(String[] args) throws IOException {
        String inputFilePath = "C:\\Users\\martin.asenov\\Desktop\\PROJECTS\\SEASON 4\\OE-IJH\\OE-IJH-H-22_Locked_24.04.2023.pdf_Bookmarked.pdf";
        String outputFilePath = inputFilePath+"_Duplicates Removed.pdf";

        int duplicatePagesFound = 0;

        try (PDDocument document = PDDocument.load(new File(inputFilePath))) {
            PDFTextStripper textStripper = new PDFTextStripper();
            PDDocumentOutline outline = document.getDocumentCatalog().getDocumentOutline();

            List<PageRange> bookmarkRanges = getBookmarkRanges(outline.getFirstChild(), document);

            for (PageRange range : bookmarkRanges) {
                List<String> pageContents = new ArrayList<>();

                for (int pageIndex = range.start; pageIndex <= range.end; pageIndex++) {
                    System.out.println("Processing page: " + (pageIndex + 1));
                    textStripper.setStartPage(pageIndex + 1);
                    textStripper.setEndPage(pageIndex + 1);
                    String pageText = textStripper.getText(document);

                    if (!isDuplicate(pageContents, pageText)) {
                        pageContents.add(pageText);
                    } else {
                        System.out.println("Duplicate page found: " + (pageIndex + 1));
                        document.removePage(pageIndex);
                        pageIndex--; // Decrement pageIndex to account for the removed page
                        range.end--; // Decrement the end index of the range
                        duplicatePagesFound++;
                    }
                }
            }

            document.save(outputFilePath);
        }

        System.out.println("Duplicate pages found: " + duplicatePagesFound);
        System.out.println("Pages removed: " + duplicatePagesFound);
    }

    private static List<PageRange> getBookmarkRanges(PDOutlineItem outlineItem, PDDocument document) throws IOException {
        List<PageRange> ranges = new ArrayList<>();

        while (outlineItem != null) {
            System.out.println("Processing outline item: " + outlineItem.getTitle());

            PDPage startPage = null;
            PDPage endPage = null;
            PDPageTree pageTree = document.getPages();
            PDDestination destination = outlineItem.getDestination();

            if (destination == null && outlineItem.getAction() instanceof PDActionGoTo action) {
                destination = action.getDestination();
            }

            if (destination instanceof PDPageDestination) {
                startPage = ((PDPageDestination) destination).getPage();
            } else if (destination instanceof PDNamedDestination) {
                PDPageDestination pageDestination = document.getDocumentCatalog().findNamedDestinationPage((PDNamedDestination) destination);
                if (pageDestination != null) {
                    startPage = pageDestination.getPage();
                }
            }

            if (outlineItem.getNextSibling() != null) {
                destination = outlineItem.getNextSibling().getDestination();
                if (destination == null && outlineItem.getNextSibling().getAction() instanceof PDActionGoTo action) {
                    destination = action.getDestination();
                }

                if (destination instanceof PDPageDestination) {
                    endPage = ((PDPageDestination) destination).getPage();
                } else if (destination instanceof PDNamedDestination) {
                    PDPageDestination pageDestination = document.getDocumentCatalog().findNamedDestinationPage((PDNamedDestination) destination);
                    if (pageDestination != null) {
                        endPage = pageDestination.getPage();
                    }
                }
            }

            if (startPage != null && endPage != null) {
                int startIndex = pageTree.indexOf(startPage);
                int endIndex = pageTree.indexOf(endPage) - 1;
                System.out.println("Bookmark range: " + startIndex + " - " + endIndex);
                ranges.add(new PageRange(startIndex, endIndex));
            } else {
                System.out.println("Start or end page not found");
            }

            if (outlineItem.hasChildren()) {
                ranges.addAll(getBookmarkRanges(outlineItem.getFirstChild(), document));
            }

            outlineItem = outlineItem.getNextSibling();
        }

        return ranges;
    }

    private static boolean isDuplicate(List<String> pageContents, String pageText) {
        return pageContents.stream().anyMatch(content -> content.trim().equalsIgnoreCase(pageText.trim()));
    }

    private static class PageRange {
        int start;
        int end;

        public PageRange(int start, int end) {
            this.start = start;
            this.end = end;
        }
    }
}