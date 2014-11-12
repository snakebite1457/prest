package pe.meyer;

import com.itextpdf.text.BaseColor;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Rectangle;
import com.itextpdf.text.pdf.*;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.io.*;
import java.util.*;
import java.util.stream.IntStream;

/**
 * Created by Dennis Meyer on 11/12/2014.
 */
public class UpdateToolTip {
    private static boolean debugMode = false;

    private static File pdfPathOld;
    private static File pdfPathNew;
    private static File excelPath;
    private static int sheetNumber;
    private static int colPoints;
    private static List<Integer> colsInfo;
    private static HashMap<String, String> pointsWithInformation;

    private static List<String> copyFields;

    public static Properties properties;

    public static void main(String[] args) {
        colsInfo = new ArrayList<>();
        pointsWithInformation = new HashMap<>();
        copyFields = new ArrayList<>();

        try {
            System.out.println("Verarbeitung starten");
            System.out.println("Konfigurationsdatei auslesen");
            readConfigFile();
            System.out.println("Exceldatei auslesen");
            readExcelFile();
            System.out.println("PDF erstellen");
            updateToolTips();
            System.out.println("Fertig");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void readExcelFile() {
        try {
            HashMap<Integer, String> headLines = new HashMap<>();

            FileInputStream fileInputStream = new FileInputStream(excelPath);
            HSSFWorkbook workbook = new HSSFWorkbook(fileInputStream);
            HSSFSheet sheet = workbook.getSheetAt(sheetNumber);

            Iterator<Row> rowIterator = sheet.iterator();
            while (rowIterator.hasNext())
            {
                Row row = rowIterator.next();
                Iterator<Cell> cellIterator = row.cellIterator();

                String point = "";
                StringBuilder stringBuilder = new StringBuilder();
                while (cellIterator.hasNext())
                {
                    Cell cell = cellIterator.next();

                    if (row.getRowNum() == 0) {
                        switch (cell.getCellType())
                        {
                            case Cell.CELL_TYPE_NUMERIC:
                                headLines.put(cell.getColumnIndex(), Integer
                                        .toString((int)cell.getNumericCellValue()));
                                break;
                            case Cell.CELL_TYPE_STRING:
                                headLines.put(cell.getColumnIndex(), cell
                                        .getStringCellValue());
                                break;
                        }
                        continue;
                    }

                    if (cell.getColumnIndex() == colPoints) {
                        point = Integer.toString((int)cell.getNumericCellValue());
                        stringBuilder.append("Nummer: " + point);
                        stringBuilder.append(System.getProperty("line" +
                                ".separator"));
                        continue;
                    }

                    if (colsInfo.contains(cell.getColumnIndex())) {
                        String value = "";
                        switch (cell.getCellType())
                        {
                            case Cell.CELL_TYPE_NUMERIC:
                                value = Integer.toString((int) cell
                                        .getNumericCellValue());
                                break;
                            case Cell.CELL_TYPE_STRING:
                                value = cell.getStringCellValue();
                                break;
                        }

                        stringBuilder.append(headLines.get(cell.getColumnIndex()));
                        stringBuilder.append( ": ");
                        stringBuilder.append(value);
                        stringBuilder.append(System.getProperty("line.separator"));
                    }
                }
                pointsWithInformation.put(point, stringBuilder.toString());
            }
            fileInputStream.close();

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void updateToolTips() {
        try {
            PdfReader pdfReader = new PdfReader(pdfPathOld.getAbsolutePath());

            PdfStamper stamper = new PdfStamper(pdfReader
                    , new FileOutputStream(pdfPathNew));

            AcroFields pdfForm = pdfReader.getAcroFields();
            Map<String, AcroFields.Item> itemMap = pdfForm.getFields();

            List<String> keys = new ArrayList<>();
            stamper.getAcroFields().getFields().forEach(
                    (key, item) -> keys.add(key));
            keys.stream()
                    .filter(key -> !copyFields.contains(key))
                    .forEach(key -> stamper.getAcroFields().removeField(key));


            for(String key : itemMap.keySet()) {
                List<AcroFields.FieldPosition> positions = pdfForm
                        .getFieldPositions(key);

                Rectangle rect = positions.get(0).position;

                if (copyFields.contains(key)) {
                    continue;
                }

                switch (pdfForm.getFieldType(key)) {
                    case AcroFields.FIELD_TYPE_TEXT:
                        TextField textField = new TextField(stamper
                                .getWriter(), rect, key);

                        textField.setVisibility(TextField.VISIBLE);
                        textField.setBackgroundColor(BaseColor.BLUE);

                        PdfFormField pdfFormField = textField.getTextField();
                        if (pointsWithInformation.containsKey(key)) {
                            pdfFormField.setUserName(pointsWithInformation.get(key));
                        }
                        else {
                            pdfFormField.setUserName("Nummer: " + key);
                        }
                        stamper.addAnnotation(pdfFormField, 1);
                }
            }
            stamper.close();


        } catch (IOException e) {
            e.printStackTrace();

        } catch (DocumentException e) {
            e.printStackTrace();
        }
    }

    private static void readConfigFile() throws IOException {
        String currentLocation = new File(UpdateToolTip.class.getProtectionDomain()
                .getCodeSource().getLocation().getPath()).getParent();

        properties = new Properties();
        properties.loadFromXML(debugMode
                ? UpdateToolTip.class.getResourceAsStream("properties.xml")
                : new FileInputStream(currentLocation +
                File.separator + "properties.xml"));

        pdfPathOld = new File(properties.getProperty("pdfPathOld"));
        pdfPathNew = new File(properties.getProperty("pdfPathNew"));
        excelPath = new File(properties.getProperty("excelPath"));
        sheetNumber = Integer.parseInt(properties.getProperty("sheetNumber"));
        colPoints = Integer.parseInt(properties.getProperty("columnPointNumbers"));

        final String[] oCols = properties.getProperty("columnOtherInformation").split(",");
        IntStream.range(0, oCols.length)
                .forEach(a -> colsInfo.add(a,Integer.valueOf(oCols[a])));

        final String[] cFields = properties.getProperty("pdfCopyFields").split(",");
        IntStream.range(0, cFields.length)
                .forEach(a -> copyFields.add(a, cFields[a]));

        if (pdfPathNew.exists()) {
            pdfPathNew.delete();
        }

        pdfPathNew.createNewFile();
    }
}
