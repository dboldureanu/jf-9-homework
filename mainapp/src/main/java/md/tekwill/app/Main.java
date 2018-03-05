package md.tekwill.app;

import md.tekwill.GenerateDocument;

public class Main {
    public static void main(String... args) throws Exception {
        GenerateDocument generateDocument = new GenerateDocument();
        generateDocument.generateDocument("Benjamin Sisko", "BenjaminSisko.jpg", "Output.docx");
    }
}
