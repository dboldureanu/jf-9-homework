package md.tekwill;

import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.*;

import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

public class GenerateDocument {

    public void generateDocument(String yourName, String imageFileName, String filePath) throws Exception {
        try (XWPFDocument doc = new XWPFDocument()) {
            XWPFParagraph p1 = doc.createParagraph();
            p1.setAlignment(ParagraphAlignment.CENTER);

            XWPFRun r1 = p1.createRun();
            r1.setBold(true);
            r1.setText("Hello, " + yourName);
            r1.setBold(true);
            r1.setFontFamily("Arial");
            r1.setFontSize(20);
            r1.setTextPosition(100);

            // ----------
            // Adding image
            // ----------
            XWPFParagraph image = doc.createParagraph();
            image.setAlignment(ParagraphAlignment.CENTER);

            XWPFRun imageRun = image.createRun();
            imageRun.setTextPosition(20);

            Path imagePath = Paths.get(ClassLoader.getSystemResource(imageFileName).toURI());
            imageRun.addPicture(Files.newInputStream(imagePath),
                    XWPFDocument.PICTURE_TYPE_PNG, imagePath.getFileName().toString(),
                    Units.toEMU(200), Units.toEMU(200));

            try (FileOutputStream out = new FileOutputStream(filePath)) {
                doc.write(out);
            }
        }
    }
}
