package Bookmark.Nomad;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDResources;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.interactive.annotation.PDAnnotationWidget;
import org.apache.pdfbox.pdmodel.interactive.form.PDAcroForm;
import org.apache.pdfbox.pdmodel.interactive.form.PDTextField;
import org.apache.pdfbox.cos.COSName;
import org.apache.pdfbox.pdmodel.font.PDType1Font;

import java.io.File;
import java.io.IOException;

public class AddTextFieldToPDF {

    public static void main(String[] args) {
        String inputFilePath = "C:\\Users\\martin.asenov\\Downloads\\PAGE 1 OF WORKORDERTASKCARD_39280475.pdf";
        String outputFilePath = "C:\\Users\\martin.asenov\\Downloads\\PAGE 1 OF WORKORDERTASKCARD_39280475_OUTPUT.pdf";

        try {
            PDDocument document = PDDocument.load(new File(inputFilePath));
            PDAcroForm acroForm = document.getDocumentCatalog().getAcroForm();

            if (acroForm == null) {
                acroForm = new PDAcroForm(document);
                document.getDocumentCatalog().setAcroForm(acroForm);
            }

            // Set the default resources for the AcroForm
            PDResources dr = new PDResources();
            acroForm.setDefaultResources(dr);

            // Adding the default font to the resources
            dr.put(COSName.getPDFName("Helv"), PDType1Font.HELVETICA);

            PDTextField textField = new PDTextField(acroForm);
            textField.setPartialName("example");

            PDAnnotationWidget widget = textField.getWidgets().get(0);
            PDRectangle rect = new PDRectangle(49, 225, 465, 223);
            widget.setRectangle(rect);


            // Set the default appearance string for the text field with navy blue color
            textField.setDefaultAppearance("/Helv 0 Tf 0 0 0.502 rg");

            //textField.setDefaultAppearance("/Helv 9 Tf 0 g");             BLACK

            // Set the text field to be multiline
            textField.setMultiline(true);


            acroForm.getFields().add(textField);
            PDPage firstPage = document.getPage(0);
            firstPage.getAnnotations().add(widget);

            document.save(new File(outputFilePath));
            document.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}