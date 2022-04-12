import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.parsers.SAXParser;
import javax.xml.parsers.SAXParserFactory;
import java.io.*;
import java.net.URISyntaxException;
import java.nio.file.Path;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

public class LecturesScannerV1 {

    //Reference file name
    final String REFERENCE_FILE_NAME = "reference.docx";
    //Path to the directory where the files to be scanned are located
    private Path reference_path;
    //Reference file path
    private File reference_file;
    //Docx scan files quantity
    private int scan_files_quantity;

    ///Scan stuff
    private List<String> referenceListYellow = new ArrayList<>();
    private List<String> referenceListCyan = new ArrayList<>();

    private ArrayList<String> listUnderTestYellow = new ArrayList<>();
    private ArrayList<String> listUnderTestCyan = new ArrayList<>();
    private String studentName = "####";

    //SET isRefListParsingNow = false WHEN PARSING scanDocument.xml !!!!!!!
    private boolean isRefListParsingNow = false;

    //REPORT STRINGS ARRAY
    private ArrayList<String> report;

    public LecturesScannerV1() {
        getParentDirectoryPath();

        report = new ArrayList<>();
        report.add("Отчет о проверке работ");
        report.add("№ | Фамилия | % совп. желтого | % совп. бирюзового");

        mainFunction();
    }

    private void mainFunction() {
        File referenceDirectory = new File(String.valueOf(reference_path));

        //1# Get files list in reference directory
        File[] filesList = referenceDirectory.listFiles();

        //2# If reference file was not found, terminate program
        if (!refFileCheck(filesList)) {
            return;
        }

        //3# Extracting referenceDocument.xml
        reference_file = new File(reference_path + File.separator + REFERENCE_FILE_NAME);
        extractDocXml(reference_file, "reference");

        scan_files_quantity = filesList.length;

        //4# Files list optimization
        for (int i = 0; i < filesList.length; i++) {
            //if file is file and file isn`t .docx
            if (!isDocx(filesList[i]) || filesList[i].getName().equals(REFERENCE_FILE_NAME)) {
                filesList[i] = null;
                scan_files_quantity--;
            }

        }

        //5# If there are no scan files, terminate program
        if (scan_files_quantity == 0) {
            File file = new File(reference_path.toString() + "\\SCAN_FILES_NOT_FOUND.txt");
            try {
                file.createNewFile();
            } catch (IOException e) {
                e.printStackTrace();
            }

            cleanTrashXML();
            return;
        }

        //6# Container for students marks
        ArrayList<String> studentsMarks = new ArrayList<>();

        //7# Scanning referenceDocument.xml
        isRefListParsingNow = true;
        parseXML();
        System.out.println("Reference data:");
        System.out.println("Yellow: " + referenceListYellow);
        System.out.println("Cyan: " + referenceListCyan);

        //8# Scanning scanDocument.xml
        isRefListParsingNow = false;

        int report_lines_counter = 1;
        for (int i = 0; i < filesList.length; i++) {
            if (filesList[i] != null) {
                extractDocXml(filesList[i], "scan");

                //SET isRefListParsingNow = false WHEN PARSING scanDocument.xml !!!!!!!
                parseXML();

                //Creating new line of report
                StringBuilder report_line = new StringBuilder();
                report_line.append(report_lines_counter++).append(") ");

                report_line.append(studentName).append("\t");
                studentName = "####";

                //A pattern for rounding to two decimal places
                DecimalFormat decimalFormat = new DecimalFormat("#.##");

                //Calculation the percentage of matches of highlighted YELLOW text
                report_line
                        .append(decimalFormat.format(compareHighLightedYellowText()))
                        .append("\t\t");

                //Calculation the percentage of matches of highlighted CYAN text
                report_line
                        .append(decimalFormat.format(compareHighLightedCyanText()))
                        .append("\t\t");

                printReferenceYellow();
                printScanYellow();

                //Adding new line to report
                report.add(report_line.toString());

                //Deleting path of docx that was scanned
                filesList[i] = null;

                //Renewing lists
                listUnderTestYellow = new ArrayList<>();
                listUnderTestCyan = new ArrayList<>();
            }
        }

        //9# Printing report in report.txt
        printReport();

        //10# DELETER TRASH XML FILES
        cleanTrashXML();

    }

    private void getParentDirectoryPath() {
        try {
            String jarPath = getClass()
                    .getProtectionDomain()
                    .getCodeSource()
                    .getLocation()
                    .toURI()
                    .getPath();

            jarPath = jarPath.substring(1);
            reference_path = Path.of(jarPath).getParent();

        } catch (URISyntaxException e) {
            e.printStackTrace();
        }
    }

    private boolean refFileCheck(File[] filesList) {
        boolean result = false;
        for (File file : filesList) {
            if (file.getName().equals(REFERENCE_FILE_NAME)) result = true;
        }

        if (result) {
            return result;
        } else {
            File file = new File(reference_path.toString() + "\\REFERENCE_FILE_NOT_FOUND.txt");

            try {
                file.createNewFile();
            } catch (IOException e) {
                e.printStackTrace();
            }

            return result;
        }
    }

    private boolean isDocx(File file) {
        String name = file.getName();
        try {
            name = name.substring(file.getName().length() - 5);
        } catch (StringIndexOutOfBoundsException e) {
        }
        if (name.equals(".docx")) return true;
        return false;
    }

    private void extractDocXml(File file, String additionalString) {

        try (ZipInputStream zin = new ZipInputStream(new FileInputStream(file.getPath()))) {
            ZipEntry entry;
            String name;
            while ((entry = zin.getNextEntry()) != null) {
                name = entry.getName(); // получим название файла
                // распаковка
                if (name.equals("word/document.xml")) {
                    FileOutputStream fout = new FileOutputStream
                            (file.getParent() + "\\" + additionalString + "Document.xml");
                    for (int c = zin.read(); c != -1; c = zin.read()) {
                        fout.write(c);
                    }
                    fout.flush();
                    zin.closeEntry();
                    fout.close();
                }
            }
        } catch (Exception ex) {
            System.out.println(ex.getMessage());
        }
    }

    private void printReport() {
        try (FileWriter writer = new FileWriter(reference_path + "\\report.txt", false)) {
            // запись всей строки
            for (String text : report) {
                writer.write(text + "\n");
            }
            writer.flush();
        } catch (IOException ex) {

            System.out.println(ex.getMessage());
        }
    }

    private void cleanTrashXML() {
        File file = new File(reference_path + "\\referenceDocument.xml");
        file.delete();
        file = new File(reference_path + "\\scanDocument.xml");
        file.delete();

    }

    ////////////////////////////////////////////////////////////////
    /// SCANNER STUFF///
    ///////////////////////////////////////////////////////////////

    class XMLHandler extends DefaultHandler {

        private boolean yellowHighlighterFlag = false, cyanHighlighterFlag = false,
                redHighlighterFlag = false;

        private String lastElementName;

        @Override
        public void startElement(String uri, String localName, String qName, Attributes attributes) throws SAXException {
            lastElementName = qName;

            if (qName.equals("w:highlight") && attributes.getValue("w:val").equals("yellow")) {
                yellowHighlighterFlag = true;
            }
            if (qName.equals("w:highlight") && attributes.getValue("w:val").equals("cyan")) {
                cyanHighlighterFlag = true;
            }
            if (qName.equals("w:highlight") && attributes.getValue("w:val").equals("red")) {
                redHighlighterFlag = true;
            }
        }

        @Override
        public void characters(char[] ch, int start, int length) throws SAXException {
            String information = new String(ch, start, length);
            information = information.replace("\n", " ").trim();

            //INSERT HIGHLIGHTED YELLOW TEXT IN CERTAIN LIST
            if (!information.isEmpty() && lastElementName.equals("w:t") && yellowHighlighterFlag) {
                if (isRefListParsingNow) {
                    referenceListYellow.add(information);
                } else {
                    listUnderTestYellow.add(information);
                }
                yellowHighlighterFlag = false;
            }

            //INSERT HIGHLIGHTED CYAN TEXT IN CERTAIN LIST
            if (!information.isEmpty() && lastElementName.equals("w:t") && cyanHighlighterFlag) {
                if (isRefListParsingNow) {
                    referenceListCyan.add(information);
                } else {
                    listUnderTestCyan.add(information);
                }
                cyanHighlighterFlag = false;
            }

            //INSERT HIGHLIGHTED RED TEXT IN STUDENT NAME VARIABLE
            if (!information.isEmpty() && lastElementName.equals("w:t") && redHighlighterFlag) {
                studentName = information;
                redHighlighterFlag = false;
            }
        }
    }

    private void parseXML() {
        SAXParserFactory factory = SAXParserFactory.newInstance();
        SAXParser parser = null;
        try {
            parser = factory.newSAXParser();
        } catch (ParserConfigurationException | SAXException e) {
        }


        XMLHandler handler = new XMLHandler();
        try {
            if (isRefListParsingNow) {
                parser.parse(new File(reference_path + "\\referenceDocument.xml"), handler);
            } else {
                parser.parse(new File(reference_path + "\\scanDocument.xml"), handler);
            }

        } catch (SAXException | IOException ex) {
        }
    }

    private double compareHighLightedYellowText() {
        //YELLOW
        int referenceListYellowWeight = 0;
        for (String string : referenceListYellow) {
            referenceListYellowWeight += string.length();
        }

        int scanListYellowWeight = 0;
        int scan_init_position = 0;
        for (int i = 0; i < listUnderTestYellow.size(); i++) {
            for (int j = scan_init_position; j < referenceListYellow.size(); j++) {
                if (listUnderTestYellow.get(i).equals(referenceListYellow.get(j))) {

                    System.out.println(listUnderTestYellow.get(i));
                    System.out.println("Weight = " + listUnderTestYellow.get(i).length());
                    System.out.println();

                    scan_init_position = j;
                    scanListYellowWeight += listUnderTestYellow.get(i).length();
                    break;
                }
            }
        }

        System.out.println();
        System.out.println("Reference Yellow Weight = " + referenceListYellowWeight);
        System.out.println("Scan Yellow Weight = " + scanListYellowWeight);

        return (double) scanListYellowWeight * 100.0 / (double) referenceListYellowWeight;
    }

    private double compareHighLightedCyanText() {
        //CYAN
        int referenceListCyanWeight = 0;
        for (String string : referenceListCyan) {
            referenceListCyanWeight += string.length();
        }
        int scanListCyanWeight = 0;
        int scan_init_position = 0;
        for (int i = 0; i < listUnderTestCyan.size(); i++) {
            for (int j = scan_init_position; j < referenceListCyan.size(); j++) {
                if (listUnderTestCyan.get(i).equals(referenceListCyan.get(j))) {
                    scan_init_position = j;
                    scanListCyanWeight += listUnderTestCyan.get(i).length();
                    break;
                }
            }
        }
        return (double) scanListCyanWeight * 100.0 / (double) referenceListCyanWeight;
    }

    private void printReferenceYellow() {
        try (FileWriter writer = new FileWriter(reference_path + "\\referenceYellow.txt", false)) {
            // запись всей строки
            for (String text : referenceListYellow) {
                writer.write(text + "\n");
            }
            writer.flush();
        } catch (IOException ex) {
            System.out.println(ex.getMessage());
        }
    }

    private void printScanYellow() {
        try (FileWriter writer = new FileWriter(reference_path + "\\scanYellow.txt", false)) {
            // запись всей строки
            for (String text : listUnderTestYellow) {
                writer.write(text + "\n");
            }
            writer.flush();
        } catch (IOException ex) {
            System.out.println(ex.getMessage());
        }
    }


}
