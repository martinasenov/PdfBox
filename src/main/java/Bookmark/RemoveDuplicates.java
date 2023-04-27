package Bookmark;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.interactive.documentnavigation.outline.PDOutlineItem;
import org.apache.pdfbox.text.PDFTextStripper;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class RemoveDuplicates {
    public static void main(String[] args) throws IOException {
        String inputFilePath = "C:\\Users\\martin.asenov\\Desktop\\PROJECTS\\SEASON 4\\OE-IJH\\OE-IJH-H-22_Locked_24.04.2023.pdf_Bookmarked.pdf";
        String outputFilePath = inputFilePath+"_Duplicates Removed.pdf";

        try (PDDocument document = PDDocument.load(new File(inputFilePath))) {
            PDOutlineItem root = document.getDocumentCatalog().getDocumentOutline().getFirstChild();

            Map<Integer, String> pagesText = new HashMap<>();
            List<Integer> pagesToRemove = new ArrayList<>();

            while (root != null) {
                int[] range = getBookmarkRanges(root, document);

                if (range != null) {
                    System.out.printf("Processing bookmark range: %d - %d%n", range[0], range[1]);

                    boolean isFirstPage = true;
                    for (int i = range[0]; i <= range[1]; i++) {
                        String currentPageText = extractText(document, document.getPage(i - 1));

                        if (pagesText.containsValue(currentPageText) && !isFirstPage) {
                            pagesToRemove.add(i);
                        } else {
                            pagesText.put(i, currentPageText);
                        }
                        isFirstPage = false;
                    }
                    // Clear the pagesText map after processing the bookmark range
                    pagesText.clear();
                }

                root = root.getNextSibling();
            }

            System.out.println("Duplicate pages found: " + pagesToRemove.size());

            for (int i = pagesToRemove.size() - 1; i >= 0; i--) {
                document.removePage(pagesToRemove.get(i) - 1);
            }

            System.out.println("Pages removed: " + pagesToRemove.size());

            document.save(new File(outputFilePath));
        }
    }

    private static int[] getBookmarkRanges(PDOutlineItem outlineItem, PDDocument document) {
        int[] range = new int[2];

        try {
            PDPage startPage = outlineItem.findDestinationPage(document);
            PDOutlineItem nextSibling = outlineItem.getNextSibling();
            PDPage endPage = nextSibling != null ? nextSibling.findDestinationPage(document) : document.getPage(document.getNumberOfPages() - 1);

            range[0] = document.getPages().indexOf(startPage) + 1;
            range[1] = document.getPages().indexOf(endPage) + 1;

            if (range[0] == -1 || range[1] == -1) {
                throw new IOException("Start or end page not found");
            }

        } catch (IOException e) {
            System.out.println("Start or end page not found for outline item: " + outlineItem.getTitle());
            return null;
        }

        return range;
    }

    private static String extractText(PDDocument document, PDPage page) throws IOException {
        PDFTextStripper textStripper = new PDFTextStripper();
        textStripper.setStartPage(document.getPages().indexOf(page) + 1);
        textStripper.setEndPage(document.getPages().indexOf(page) + 1);

        return textStripper.getText(document);
    }

}