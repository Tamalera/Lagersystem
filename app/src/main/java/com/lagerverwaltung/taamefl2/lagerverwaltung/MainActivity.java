package com.lagerverwaltung.taamefl2.lagerverwaltung;

import android.content.Intent;
import android.content.res.AssetManager;
import android.support.v7.app.AppCompatActivity;
import android.os.Bundle;
import android.util.Log;
import android.view.View;
import android.widget.Button;

import com.google.zxing.client.android.Intents;
import com.google.zxing.integration.android.IntentIntegrator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.InputStream;
import java.util.Objects;

public class MainActivity extends AppCompatActivity {

    private CellAddress address;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        System.setProperty("org.apache.poi.javax.xml.stream.XMLInputFactory", "com.fasterxml.aalto.stax.InputFactoryImpl");
        System.setProperty("org.apache.poi.javax.xml.stream.XMLOutputFactory", "com.fasterxml.aalto.stax.OutputFactoryImpl");
        System.setProperty("org.apache.poi.javax.xml.stream.XMLEventFactory", "com.fasterxml.aalto.stax.EventFactoryImpl");
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        final Button addWare = findViewById(R.id.addBtn);
        addWare.setOnClickListener(new View.OnClickListener() {
            public void onClick(View v) {
                scanQRCode();
            }
        });
    }

    private void scanQRCode() {
        IntentIntegrator integrator = new IntentIntegrator(this);
        integrator.setCaptureActivity(MyCaptureActivity.class);
        integrator.setDesiredBarcodeFormats(IntentIntegrator.QR_CODE_TYPES);
        integrator.setOrientationLocked(false);
        integrator.addExtra(Intents.Scan.BARCODE_IMAGE_ENABLED, true);
        integrator.initiateScan();
    }

    @Override
    protected void onActivityResult(int requestCode, int resultCode, Intent data) {
        if (requestCode == IntentIntegrator.REQUEST_CODE
                && resultCode == RESULT_OK) {
            Bundle extras = data.getExtras();
            if (extras != null) {
                saveQRCode(Objects.requireNonNull(extras.getString(Intents.Scan.RESULT)));
                augmentWare();
            }
        }
    }

    private void saveQRCode(String scannedName) {
        searchWareInExcel(scannedName);
    }

    private void searchWareInExcel(String scannedName) {
        try {
            InputStream myInput;
            // initialize asset manager
            AssetManager assetManager = getAssets();
            //  open excel sheet
            myInput = assetManager.open("inventar.xlsx");
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

    private void augmentWare() {
        try {
            InputStream myInput;
            // initialize asset manager
            AssetManager assetManager = getAssets();
            //  open excel sheet
            myInput = assetManager.open("inventar.xlsx");
            // Finds the workbook instance for XLSX file
            XSSFWorkbook myWorkBook = new XSSFWorkbook (myInput);
            // Return first sheet from the XLSX workbook
            XSSFSheet mySheet = myWorkBook.getSheetAt(0);

            Cell cellToChange = mySheet.getRow(this.address.getRow()).getCell(this.address.getColumn() + 1);
            cellToChange.setCellValue(cellToChange.getNumericCellValue() + 1);

        } catch (Exception e) {
            Log.e("ERROR: ", "error "+ e.toString());
        }
    }
}
