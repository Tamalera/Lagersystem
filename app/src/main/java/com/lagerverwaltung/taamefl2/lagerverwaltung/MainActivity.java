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

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;
import java.util.Objects;

public class MainActivity extends AppCompatActivity {

    private String buttonCode;
    private ExcelChanges myExcel = new ExcelChanges(this);
    private ShoppingList shoppingList = new ShoppingList(this);

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
                setButtonCode("add");
                scanQRCode();
            }
        });

        final Button removeWare = findViewById(R.id.removeBtn);
        removeWare.setOnClickListener(new View.OnClickListener() {
            public void onClick(View v) {
                setButtonCode("remove");
                scanQRCode();
            }
        });

        final Button previewShoppingList = findViewById(R.id.shoppingListBtn);
        previewShoppingList.setOnClickListener(new View.OnClickListener() {
            public void onClick(View v) {
                showPreview(shoppingList.getPreview());
            }
        });

        final Button addNewEntry = findViewById(R.id.newBtn);
        addNewEntry.setOnClickListener(new View.OnClickListener() {
            public void onClick(View v) {
                myExcel.addNewEntry();
                showExcel();
            }
        });

        final Button addNewExcel = findViewById(R.id.addExcel);
        addNewExcel.setOnClickListener(new View.OnClickListener() {
            public void onClick(View v) {
                CopyAssets();
                showExcel();
            }
        });
    }

    private void showPreview(List<String> preview) {
        TextView previewField = findViewById(R.id.previewField);
        previewField.setText("");
        for (String element: preview) {
            previewField.append("\n" + element);
        }
    }

    private void setButtonCode(String code) {
        this.buttonCode = code;
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
                myExcel.saveQRCode(Objects.requireNonNull(extras.getString(Intents.Scan.RESULT)));
                if (this.buttonCode.equals("add")){
                    myExcel.augmentWare();
                } else if (this.buttonCode.equals("remove")){
                    myExcel.takeWare();
                }
            }
        }
    }

    private void CopyAssets() {
        AssetManager assetManager = getAssets();
        InputStream in;
        OutputStream out;
        try {
            in = assetManager.open("inventar.xlsx");
            out = new FileOutputStream(getFilesDir() + "/inventar.xlsx");
            copyFile(in, out);
            in.close();
            out.flush();
            out.close();
        } catch(Exception e) {
            Log.e("tag", e.getMessage());
        }
    }
    private void copyFile(InputStream in, OutputStream out) throws IOException {
        byte[] buffer = new byte[1024];
        int read;
        while((read = in.read(buffer)) != -1){
            out.write(buffer, 0, read);
        }
    }

    private void showExcel(){
        try {
            FileInputStream myInput;
            myInput = openFileInput("inventar.xlsx");
            XSSFWorkbook myWorkBook = new XSSFWorkbook (myInput);
            XSSFSheet mySheet = myWorkBook.getSheetAt(0);

            for (Row row: mySheet) {
                for (Cell myCell: row) {
                    Log.i("CELL: ", myCell.toString());
                }
            }
            myInput.close();

        } catch (Exception e) {
            Log.e("ERROR: ", "error "+ e.toString());
        }
    }
}
