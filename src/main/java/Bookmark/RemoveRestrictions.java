package Bookmark;

import org.apache.pdfbox.pdmodel.PDDocument;

import java.io.File;
import java.io.IOException;

public class RemoveRestrictions {
    public static void main(String[] args) {
        try {
            String filePath="C:\\Users\\martin.asenov\\Downloads\\WORK CARD - P4-BBJ.pdf";
            PDDocument document = PDDocument.load(new File(filePath));

            // Remove all security restrictions.
            document.setAllSecurityToBeRemoved(true);

            // Save the unprotected document.
            document.save(filePath+"_unprotected.pdf");
            document.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
