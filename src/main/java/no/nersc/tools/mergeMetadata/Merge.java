package no.nersc.tools.mergeMetadata;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.Scanner;

import javax.swing.JOptionPane;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.mozilla.universalchardet.UniversalDetector;


public class Merge {

    public static void main(String[] args) {
        String mdFileName = "Metadata Aleksander 2016.xlsx";
        String s0 = "Hva er navnet på excel-metadatafilen?";
        mdFileName = JOptionPane.showInputDialog(null, s0);
        ArrayList<String> headers = parseExcelFileForHeaders(mdFileName);
        System.out.println("Headers:");
        int i;
        int numberOfColumns = headers.size();
        String headerString = "";
        for (i = 0; i < numberOfColumns; i++) {
            headerString += i + ":'" + headers.get(i) + "'\r\n";
        }
        String s1 = headerString+"Hvilken kolonne (0-" + (i - 1) + ") inneholder selve navnet på filen:";
        String input = JOptionPane.showInputDialog(null, s1);
        System.out.println(input);
        System.out.print(s1);
        // Scanner sc = new Scanner(System.in);
        // int fileNameIndex = Integer.parseInt(sc.nextLine());
        int fileNameIndex = Integer.parseInt(input);
        System.out.println();
        System.out.println("Filnavn i filen:");
        ArrayList<String> fileNameColumn = getDataFromColumn(mdFileName, fileNameIndex);
        String fileNameColumnString = "";
        for (i = 0; i < fileNameColumn.size(); i++) {
            fileNameColumnString += i + ":'" + fileNameColumn.get(i) + "'\r\n";
        }

        String s2 = "Har filene noen ekstra extension (e.g. '.txt' eller '.csv'), skriv inn eventuelt extension (med .) for ja, la feltet være blank for nei:";
        System.out.print(s2);
        // String extension = sc.nextLine();
        System.out.println();
        String extension = JOptionPane.showInputDialog(null, s2);

        ArrayList<String> fileNames = getFileNames(extension);
        System.out.println("Fant følgende filnavn som vi nå leter etter:");
        for (i = 0; i < fileNames.size(); i++) {
            System.out.println(i + ":'" + fileNames.get(i) + "'");
        }
        for (i = 0; i < fileNames.size(); i++) {
            String fileName = fileNames.get(i);
            fileName = fileName.substring(0, fileName.length() - extension.length());
            int index = fileNameColumn.indexOf(fileName);
            if (index >= 0) {
                BufferedWriter writer = null;
                try {
                    String encoding = detectEncoding(fileName + extension);
                    System.out.println("Fant '" + fileName + "' på rad " + (index + 2));
                    ArrayList<String> rowData = getDataFromRow(mdFileName, index + 1);
                    writer = new BufferedWriter(new OutputStreamWriter(
                                    new FileOutputStream(fileName + "-merged-"
                                                    + System.currentTimeMillis() + extension),
                                    "UTF-8"));
                    // new FileWriter(fileName+"-merged"+extension, true));
                    for (int j = 0; j < numberOfColumns; j++) {
                        String data = rowData.get(j);
                        if (data.trim().isEmpty()) {
                            data = "N/A";
                        }
                        writer.append(headers.get(j) + ":" + data + "\r\n");
                    }
                    Scanner in = new Scanner(new File(fileName + extension), encoding);
                    while (in.hasNext()) {
                        writer.append(in.nextLine() + "\r\n");
                    }
                    in.close();
                } catch (IOException e) {
                    System.out.println(e.getMessage());
                    // TODO Auto-generated catch block
                    e.printStackTrace();
                } finally {
                    if (writer != null)
                        try {
                            writer.close();
                        } catch (IOException e) {
                            // TODO Auto-generated catch block
                            e.printStackTrace();
                        }
                }
            } else
                System.out.println("Fant ikke '" + fileName + "'");
        }
    }

    private static String detectEncoding(String fileName) throws IOException {
        UniversalDetector detector = new UniversalDetector(null);

        // (2)
        int nread;
        byte[] buf = new byte[4096];
        FileInputStream fis = new FileInputStream(new File(fileName));
        while ((nread = fis.read(buf)) > 0 && !detector.isDone()) {
            detector.handleData(buf, 0, nread);
        }
        fis.close();
        // (3)
        detector.dataEnd();

        // (4)
        String encoding = detector.getDetectedCharset();
        if (encoding != null) {
            System.out.println("Detected encoding = " + encoding);
        } else {
            System.out.println("No encoding detected.");
        }

        // (5)
        detector.reset();
        return encoding;

    }

    private static ArrayList<String> getDataFromRow(String fileName, int rowIndex) {
        ArrayList<String> rowData = new ArrayList<String>();
        XSSFWorkbook myWorkBook = null;
        try {
            File file = new File(fileName);
            FileInputStream fis = new FileInputStream(file);
            DataFormatter formatter = new DataFormatter();
            myWorkBook = new XSSFWorkbook(fis);
            XSSFSheet mySheet = myWorkBook.getSheetAt(0);
            Row row = mySheet.getRow(rowIndex);
            Iterator<Cell> cellIterator = row.cellIterator();
            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                rowData.add(formatter.formatCellValue(cell));
            }
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (myWorkBook != null)
                try {
                    myWorkBook.close();
                } catch (IOException e) {
                    // TODO Auto-generated catch block
                    e.printStackTrace();
                }
        }
        return rowData;
    }

    private static ArrayList<String> getDataFromColumn(String fileName, int fileNameIndex) {
        ArrayList<String> column = new ArrayList<String>();
        XSSFWorkbook myWorkBook = null;
        try {
            File file = new File(fileName);
            FileInputStream fis = new FileInputStream(file);
            DataFormatter formatter = new DataFormatter();
            myWorkBook = new XSSFWorkbook(fis);
            XSSFSheet mySheet = myWorkBook.getSheetAt(0);
            Iterator<Row> rowIterator = mySheet.iterator();
            Row row = rowIterator.next();
            while (rowIterator.hasNext()) {
                row = rowIterator.next();
                Cell cell = row.getCell(fileNameIndex);
                column.add(formatter.formatCellValue(cell));
            }
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (myWorkBook != null)
                try {
                    myWorkBook.close();
                } catch (IOException e) {
                    // TODO Auto-generated catch block
                    e.printStackTrace();
                }
        }
        return column;
    }

    private static ArrayList<String> getFileNames(String extension) {
        ArrayList<String> fileNames = new ArrayList<String>();
        File dir = new File(".");
        File[] directoryListing = dir.listFiles();
        if (directoryListing != null) {
            for (File child : directoryListing) {
                String fileName = child.getName();
                if (fileName.endsWith(extension))
                    fileNames.add(fileName);
            }
        }
        return fileNames;
    }

    private static ArrayList<String> parseExcelFileForHeaders(String fileName) {
        ArrayList<String> headers = null;
        XSSFWorkbook myWorkBook = null;
        try {
            headers = new ArrayList<String>();
            File file = new File(fileName);
            FileInputStream fis = new FileInputStream(file);
            DataFormatter formatter = new DataFormatter();
            myWorkBook = new XSSFWorkbook(fis);
            XSSFSheet mySheet = myWorkBook.getSheetAt(0);
            Iterator<Row> rowIterator = mySheet.iterator();
            // Parse headers + 1 row
            Row headerRow = rowIterator.next();
            Iterator<Cell> headerIterator = headerRow.cellIterator();
            while (headerIterator.hasNext()) {
                Cell headerCell = headerIterator.next();
                String header = formatter.formatCellValue(headerCell);
                headers.add(header);
            }
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (myWorkBook != null)
                try {
                    myWorkBook.close();
                } catch (IOException e) {
                    // TODO Auto-generated catch block
                    e.printStackTrace();
                }
        }
        return headers;
    }

}
