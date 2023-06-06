package Bookmark;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;

public class AirbusXMLCreator {

    public static void main(String[] args) throws IOException {


        ArrayList<String> AMMRefs = new ArrayList<>();

        String excelPath = "C:\\Users\\mitha\\IdeaProjects\\PdfBox\\src\\main\\java\\Bookmark\\AMMRefs.xlsx";
        XSSFWorkbook workbook = new XSSFWorkbook(excelPath);
        XSSFSheet sheet = workbook.getSheet("Sheet1");

        int rowCount = sheet.getPhysicalNumberOfRows();


        for (int i = 1; i < rowCount; i++) {

            String inputFilePath = sheet.getRow(i).getCell(3).getStringCellValue();

            AMMRefs.add(inputFilePath);
        }


        try {
            DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
            DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
            Document doc = dBuilder.newDocument();

            // root element
            Element rootElement = doc.createElement("JobCardBasket");
            doc.appendChild(rootElement);

            // name element
            Element name = doc.createElement("name");
            rootElement.appendChild(name);

            // orientation element
            Element orientation = doc.createElement("orientation");
            orientation.appendChild(doc.createTextNode("portrait"));
            rootElement.appendChild(orientation);

            // format element
            Element format = doc.createElement("format");
            format.appendChild(doc.createTextNode("A4"));
            rootElement.appendChild(format);

            // acType element
            Element acType = doc.createElement("acType");
            acType.appendChild(doc.createTextNode("A318,A319,A320,A321"));
            rootElement.appendChild(acType);

            // customization element
            Element customization = doc.createElement("customization");
            customization.appendChild(doc.createTextNode("69X"));
            rootElement.appendChild(customization);

            // effectivity element
            Element effectivity = doc.createElement("effectivity");
            effectivity.appendChild(doc.createTextNode("N2119"));
            rootElement.appendChild(effectivity);

            // Tasks element
            Element tasks = doc.createElement("Tasks");
            rootElement.appendChild(tasks);

            // here you can replace with your variable value
            for (String AmmRef:AMMRefs) {

                System.out.println("Processing AMM reference: " + AmmRef);

                // Task elements
            Element task1 = doc.createElement("Task");
            Element dmId1 = doc.createElement("dmId");

            if(AmmRef.length()<=12){
                dmId1.appendChild(doc.createTextNode("609898_SGML_C_EN" + AmmRef+"00"));
            }else{
                dmId1.appendChild(doc.createTextNode("609898_SGML_C_EN" + AmmRef));
            }
            task1.appendChild(dmId1);
            Element productKey1 = doc.createElement("productKey");
            productKey1.appendChild(doc.createTextNode("[N]##69X#AMM###"));
            task1.appendChild(productKey1);
            Element doctype1 = doc.createElement("doctype");
            doctype1.appendChild(doc.createTextNode("AMM"));
            task1.appendChild(doctype1);
            tasks.appendChild(task1);
            }

            // create the XML file
            TransformerFactory transformerFactory = TransformerFactory.newInstance();
            Transformer transformer = transformerFactory.newTransformer();
            DOMSource domSource = new DOMSource(doc);
            StreamResult streamResult = new StreamResult(new File("C:\\Users\\mitha\\OneDrive\\Desktop\\AirbusWorldXML\\airbusTask.xml"));

            transformer.transform(domSource, streamResult);

            System.out.println("Done creating XML File");

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}