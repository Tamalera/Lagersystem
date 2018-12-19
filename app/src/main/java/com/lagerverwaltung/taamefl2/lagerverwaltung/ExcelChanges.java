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
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

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
            FileInputStream myInput = getFileInputStream();
            XSSFWorkbook myWorkBook = new XSSFWorkbook (myInput);
            XSSFSheet mySheet = myWorkBook.getSheetAt(0);

            if (this.address != null){
                Cell cellToChange = mySheet.getRow(this.address.getRow()).getCell(this.address.getColumn() + 1);
                cellToChange.setCellValue(cellToChange.getNumericCellValue() + 1);
            }
            myInput.close();
            // Now write the output to a file
            writeToExcelFile(myWorkBook);
        } catch (Exception e) {
            Log.e("ERROR: ", "error "+ e.toString());
        }
    }

    public void takeWare() {
        try {
            FileInputStream myInput = getFileInputStream();
            XSSFWorkbook myWorkBook = new XSSFWorkbook (myInput);
            XSSFSheet mySheet = myWorkBook.getSheetAt(0);

            Cell cellToChange = mySheet.getRow(this.address.getRow()).getCell(this.address.getColumn() + 1);
            if (cellToChange.getNumericCellValue() != 0){
                cellToChange.setCellValue(cellToChange.getNumericCellValue() - 1);
            }
            myInput.close();
            // Now write the output to a file
            writeToExcelFile(myWorkBook);
        } catch (Exception e) {
            Log.e("ERROR: ", "error "+ e.toString());
        }
    }

    public List<String> getAllItemsForShopping(){
        List<String> shoppingList = new ArrayList<>();
        try {
            FileInputStream myInput = getFileInputStream();
            XSSFWorkbook myWorkBook = new XSSFWorkbook (myInput);
            XSSFSheet mySheet = myWorkBook.getSheetAt(0);

            for (Row row: mySheet) {
                if (!checkIfCellTypeIsString(row.getCell(1))){
                    if ((row.getCell(1).getNumericCellValue() <= row.getCell(2).getNumericCellValue())){
                        shoppingList.add(row.getCell(0).getStringCellValue());
                    }
                }
            }
            myInput.close();
        } catch (Exception e) {
            Log.e("ERROR: ", "error "+ e.toString());
        }

        return shoppingList;
    }

    public void addNewEntry(Ware ware) {
        try {
            FileInputStream myInput = getFileInputStream();
            XSSFWorkbook myWorkBook = new XSSFWorkbook (myInput);
            XSSFSheet mySheet = myWorkBook.getSheetAt(0);

            //Save new row
            Row newRow = mySheet.createRow(mySheet.getLastRowNum() + 1);
            newRow.createCell(0).setCellValue(ware.getName());
            newRow.createCell(1).setCellValue(ware.getNumber());
            newRow.createCell(2).setCellValue(ware.getThreshold());

            myInput.close();
            writeToExcelFile(myWorkBook);
        } catch (Exception e) {
            Log.e("ERROR: ", "error "+ e.toString());
        }
    }

    private FileInputStream getFileInputStream() throws FileNotFoundException {
        FileInputStream myInput;
        myInput = context.openFileInput("inventar.xlsx");
        return myInput;
    }

    private void writeToExcelFile(XSSFWorkbook myWorkBook) throws IOException {
        FileOutputStream outFile = new FileOutputStream(new File(context.getFilesDir() + "/inventar.xlsx"));
        myWorkBook.write(outFile);
        outFile.close();
    }

    private void searchWareInExcel(String scannedName) {
        try {
            FileInputStream myInput = getFileInputStream();
            XSSFWorkbook myWorkBook = new XSSFWorkbook (myInput);
            XSSFSheet mySheet = myWorkBook.getSheetAt(0);

            // Traversing over each row of XLSX file
            for (Row myRow : mySheet) {
                for (Cell myCell : myRow) {
                    if (getCellAddress(scannedName, myCell)) break;
                }
            }
            myInput.close();
        } catch (Exception e) {
            Log.e("ERROR: ", "error "+ e.toString());
        }
    }

    private boolean getCellAddress(String scannedName, Cell myCell) {
        if(checkIfCellTypeIsString(myCell)){
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

    private boolean checkIfCellTypeIsString (Cell cell){
        return cell.getCellTypeEnum() == CellType.STRING;
    }
}
