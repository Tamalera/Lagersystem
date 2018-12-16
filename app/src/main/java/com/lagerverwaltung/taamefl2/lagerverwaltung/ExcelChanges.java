package com.lagerverwaltung.taamefl2.lagerverwaltung;

import android.content.Context;
import android.util.Log;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;

public class ExcelChanges {

    private CellAddress address;
    private Context context;

    public ExcelChanges(Context context){
        this.context = context;
    }

    public void saveQRCode(String scannedName) {
        searchWareInExcel(scannedName);
    }

    public void augmentWare() {
        try {
            FileInputStream myInput;
            //  open excel sheet
            myInput = context.openFileInput("inventar.xlsx");
            // Finds the workbook instance for XLSX file
            XSSFWorkbook myWorkBook = new XSSFWorkbook (myInput);
            // Return first sheet from the XLSX workbook
            XSSFSheet mySheet = myWorkBook.getSheetAt(0);

            Cell cellToChange = mySheet.getRow(this.address.getRow()).getCell(this.address.getColumn() + 1);
            cellToChange.setCellValue(cellToChange.getNumericCellValue() + 1);
            myInput.close();
            // Now write the output to a file
            FileOutputStream outFile = new FileOutputStream(new File(context.getFilesDir() + "/inventar.xlsx"));
            myWorkBook.write(outFile);
            outFile.close();

        } catch (Exception e) {
            Log.e("ERROR: ", "error "+ e.toString());
        }
    }

    private void searchWareInExcel(String scannedName) {
        try {
            InputStream myInput;
            //  open excel sheet
            myInput = context.openFileInput("inventar.xlsx");
            // Finds the workbook instance for XLSX file
            XSSFWorkbook myWorkBook = new XSSFWorkbook (myInput);
            // Return first sheet from the XLSX workbook
            XSSFSheet mySheet = myWorkBook.getSheetAt(0);

            // Traversing over each row of XLSX file
            for (Row myRow : mySheet) {
                for (Cell myCell : myRow) {
                    if (getCellAddress(scannedName, myCell)) break;
                }
            }
        } catch (Exception e) {
            Log.e("ERROR: ", "error "+ e.toString());
        }
    }
    private boolean getCellAddress(String scannedName, Cell myCell) {
        if(myCell.getCellTypeEnum() == CellType.STRING){
            if (scannedName.substring(7).equals(myCell.getStringCellValue())){
                saveAddress(myCell.getAddress());
                return true;
            }
        }
        return false;
    }

    private void saveAddress(CellAddress address) {
        this.address = address;
    }

}
