package com.lagerverwaltung.taamefl2.lagerverwaltung;

import android.content.Intent;
import android.content.res.AssetManager;
import android.support.v7.app.AppCompatActivity;
import android.os.Bundle;
import android.util.Log;
import android.view.View;
import android.widget.Button;
import android.widget.TextView;

import com.google.zxing.client.android.Intents;
import com.google.zxing.integration.android.IntentIntegrator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.InputStream;
import java.util.Iterator;

public class MainActivity extends AppCompatActivity {

    private String warenName;
    private TextView previewField;

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

    private void saveQRCode(String name) {
        warenName = name;
        readExcel();
    }

    private void readExcel() {
        previewField = findViewById(R.id.previewField);
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
            // Get iterator to all the rows in current sheet
            Iterator rowIterator = mySheet.iterator();
            // Traversing over each row of XLSX file
            while (rowIterator.hasNext())
            {
                Row row = (Row) rowIterator.next();
                // For each row, iterate through each columns
                Iterator cellIterator = row.cellIterator();
                while (cellIterator.hasNext()) {
                    Cell cell = (Cell) cellIterator.next();
                    switch (cell.getCellType()) {
                        case Cell.CELL_TYPE_STRING: previewField.append(cell.getStringCellValue());
                            break;
                        case Cell.CELL_TYPE_NUMERIC: previewField.append(String.valueOf(cell.getNumericCellValue()));
                            break;
                        case Cell.CELL_TYPE_BOOLEAN: previewField.append(String.valueOf(cell.getBooleanCellValue()));
                            break;
                        default : }
                }
            }
        } catch (Exception e) {
            Log.e("Stuff: ", "error "+ e.toString());
        }
    }

    @Override
    protected void onActivityResult(int requestCode, int resultCode, Intent data) {
        if (requestCode == IntentIntegrator.REQUEST_CODE
                && resultCode == RESULT_OK) {
            Bundle extras = data.getExtras();
            assert extras != null;
            saveQRCode(extras.getString(Intents.Scan.RESULT));
        }
    }


}
