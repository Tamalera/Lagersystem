package com.lagerverwaltung.taamefl2.lagerverwaltung;

import android.content.Context;

import java.util.List;

public class ShoppingList {

    private Context context;

    public ShoppingList(Context context){
        this.context = context;
    }

    public List<String> getPreview(){
        ExcelChanges excelChanges = new ExcelChanges(this.context);
        return excelChanges.getAllItemsForShopping();
    }
}
