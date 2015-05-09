package updateexcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.commons.lang3.math.NumberUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 *
 * @author Jorge.Saldivar
 * @version 0.1
 * @since MAY-08-2015
 */
public class UpdateExcel {

    static void processExcel(String fileName) {
        InputStream inp = null;
        try {
            inp = new FileInputStream(System.getProperty("user.dir") + "\\" + fileName);
            Workbook wb = WorkbookFactory.create(inp);
            Sheet sheet = wb.getSheetAt(0);

            //Iterate through each rows one by one
            Iterator<Row> rowIterator = sheet.iterator();

            int caseNumberIndex = 0;
            int notesIndex = 0;

            Row title = rowIterator.next();
            //For each row, iterate through all the columns
            Iterator<Cell> cellIterator = title.cellIterator();

            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                //Check the cell type and format accordingly
                //Case number cell
                if (cell.getStringCellValue().toUpperCase().contains("CLOSE")) {
                    caseNumberIndex = cell.getColumnIndex();
                }
                //Long description, where case number is hidden
                if (cell.getStringCellValue().toUpperCase().contains("NOTE")) {
                    notesIndex = cell.getColumnIndex();
                }

            }

            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                String caseNumber = "";
                String notes = row.getCell(notesIndex).getStringCellValue();
                notes = notes.toUpperCase().replaceAll("\\s+", "");
                int intIndex = notes.indexOf("CASE#");
                if (intIndex == - 1) {
                    caseNumber = "-1";
                } else {

                    notes = notes.substring(intIndex, intIndex + 60);

                    int enter = 0;
                    for (int i = 0; i < notes.length(); i++) {

                        if (NumberUtils.isNumber(String.valueOf(notes.charAt(i)))) {
                            caseNumber += notes.charAt(i);
                            enter++;
                        }
                        if (enter >= 7) {
                            break;
                        }

                    }

                }

                Cell cell = row.getCell(caseNumberIndex);

                if (cell == null) {
                    cell = row.createCell(caseNumberIndex);
                }

                cell.setCellType(Cell.CELL_TYPE_NUMERIC);
                cell.setCellValue(Integer.valueOf(caseNumber));

            }

            String[] splitFileName = fileName.split("\\.");
            fileName = splitFileName[0] + "_NEW_FILE." + splitFileName[1];

            try ( // Write the output to a file
                    FileOutputStream fileOut = new FileOutputStream(System.getProperty("user.dir") + "\\" + fileName)) {
                wb.write(fileOut);
            } finally {
                wb.close();
            }
        } catch (FileNotFoundException ex) {
            Logger.getLogger(UpdateExcel.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException | InvalidFormatException ex) {
            Logger.getLogger(UpdateExcel.class.getName()).log(Level.SEVERE, null, ex);
        } finally {
            if (inp != null) {
                try {
                    inp.close();
                } catch (IOException ex) {
                    Logger.getLogger(UpdateExcel.class.getName()).log(Level.SEVERE, null, ex);
                }
            }
        }

    }

    public static void main(String[] args) {
        System.out.println("Start");

        File folder = new File(System.getProperty("user.dir"));
        File[] listOfFiles = folder.listFiles();

        for (int i = 0; i < listOfFiles.length; i++) {
            if (listOfFiles[i].isFile() && listOfFiles[i].getName().contains(".xls") && !listOfFiles[i].getName().contains("NEW_FILE")) {
                String fileName = listOfFiles[i].getName();
                processExcel(fileName);
            }
        }
        
        System.out.println("Done");

    }
}
